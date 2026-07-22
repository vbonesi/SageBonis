# -*- coding: utf-8 -*-
"""Harness reutilizável para testes UNO reais (LibreOffice headless) do
ImportadorSAGE.py. Cuida do ciclo de vida completo: sobe o soffice, copia +
injeta a macro atual numa cópia DESCARTÁVEL do SageBonis.ods real (nunca
escreve no arquivo rastreado), conecta via UNO, oferece helpers de leitura/
escrita de abas, e limpa tudo ao final (documento, processo, profile, cópia).

Requer 'soffice' no PATH e o módulo 'uno' (python3-uno / vem com o LibreOffice).

Uso típico:
    from uno_harness import TesteUno
    with TesteUno() as t:
        t.chamar_macro("gerar_ied")
        linhas = t.ler_aba("LSC")

Cada instância usa uma porta TCP e um profile/cópia próprios (isolados por
tempfile), então é seguro rodar vários testes em sequência ou até em paralelo
com portas diferentes.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import time

_RAIZ_COMPLETA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ODS_REAL = os.path.join(_RAIZ_COMPLETA, "SageBonis.ods")
PY_ATUAL = os.path.join(_RAIZ_COMPLETA, "ImportadorSAGE.py")
_SYNC_MACRO = os.path.join(os.path.dirname(_RAIZ_COMPLETA), "sync_macro.py")


class TesteUno:
    def __init__(self, porta=2100, timeout_conexao=40, timeout_boot=20):
        self.porta = porta
        self.timeout_conexao = timeout_conexao
        self.timeout_boot = timeout_boot
        self.profile_dir = None
        self.copia_ods = None
        self.processo = None
        self.doc = None
        self.desktop = None
        self.ctx = None

    def __enter__(self):
        self.profile_dir = tempfile.mkdtemp(prefix="sagebonis_lo_profile_")
        fd, self.copia_ods = tempfile.mkstemp(suffix=".ods", prefix="sagebonis_teste_")
        os.close(fd)
        shutil.copy2(ODS_REAL, self.copia_ods)
        subprocess.run(
            [sys.executable, _SYNC_MACRO, "inject", "--ods", self.copia_ods,
             "--py", PY_ATUAL, "--no-backup"],
            check=True, capture_output=True,
        )
        self.processo = subprocess.Popen(
            ["soffice", "--headless", "--norestore", "--nologo",
             "--accept=socket,host=localhost,port=%d;urp;" % self.porta,
             "-env:UserInstallation=file://%s" % self.profile_dir],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        time.sleep(self.timeout_boot)
        self._conectar()
        self.doc = self._abrir_doc()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self.doc is not None:
                self.doc.close(False)
        except Exception:
            pass
        if self.processo is not None:
            self.processo.terminate()
            try:
                self.processo.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.processo.kill()
        if self.profile_dir and os.path.isdir(self.profile_dir):
            shutil.rmtree(self.profile_dir, ignore_errors=True)
        if self.copia_ods and os.path.exists(self.copia_ods):
            os.remove(self.copia_ods)
        return False

    def _conectar(self):
        import uno
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        inicio = time.time()
        ultimo_erro = None
        while time.time() - inicio < self.timeout_conexao:
            try:
                self.ctx = resolver.resolve(
                    "uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % self.porta)
                self.desktop = self.ctx.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", self.ctx)
                return
            except Exception as e:
                ultimo_erro = e
                time.sleep(0.5)
        raise RuntimeError(
            "Não conseguiu conectar ao soffice na porta %d após %ds: %s"
            % (self.porta, self.timeout_conexao, ultimo_erro))

    def _abrir_doc(self):
        from com.sun.star.beans import PropertyValue
        p = PropertyValue()
        p.Name = "Hidden"
        p.Value = True
        url = "file://" + self.copia_ods
        return self.desktop.loadComponentFromURL(url, "_blank", 0, (p,))

    def chamar_macro(self, nome_funcao, *args):
        provider = self.doc.getScriptProvider()
        script = provider.getScript(
            "vnd.sun.star.script:ImportadorSAGE.py$%s?language=Python&location=document" % nome_funcao)
        return script.invoke(args, (), ())

    def sheet_existe(self, nome):
        return self.doc.Sheets.hasByName(nome)

    def get_sheet(self, nome):
        return self.doc.Sheets.getByName(nome)

    def ler_aba(self, nome_aba):
        """Lê uma aba inteira como lista de dicts (header->valor), ignorando
        linhas totalmente vazias."""
        sheet = self.get_sheet(nome_aba)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        n_linhas = cursor.RangeAddress.EndRow + 1
        n_cols = cursor.RangeAddress.EndColumn + 1
        dados = sheet.getCellRangeByPosition(0, 0, n_cols - 1, n_linhas - 1).getDataArray()
        headers = [str(c) for c in dados[0]]
        linhas = []
        for row in dados[1:]:
            if not any(str(c).strip() for c in row):
                continue
            linhas.append({headers[i]: str(row[i]) for i in range(len(headers))})
        return linhas

    def contar_linhas(self, nome_aba):
        if not self.sheet_existe(nome_aba):
            return 0
        return len(self.ler_aba(nome_aba))

    def headers_de(self, nome_aba):
        sheet = self.get_sheet(nome_aba)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        return [sheet.getCellByPosition(c, 0).getString() for c in range(addr.EndColumn + 1)]

    def escrever_linha(self, nome_aba, linha_num, valores):
        """Escreve vários campos numa linha (0-indexed) de uma aba, casando
        por nome de cabeçalho já existente na aba."""
        sheet = self.get_sheet(nome_aba)
        headers = self.headers_de(nome_aba)
        col = {h: i for i, h in enumerate(headers)}
        for campo, valor in valores.items():
            sheet.getCellByPosition(col[campo], linha_num).setString(valor)

    def proxima_linha_livre(self, nome_aba):
        sheet = self.get_sheet(nome_aba)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        return cursor.RangeAddress.EndRow + 1
