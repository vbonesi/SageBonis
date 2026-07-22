# -*- coding: utf-8 -*-
"""Roda todos os testes da Trilha Completa em sequência: os 4 smoke tests em
memória (rápidos, sem dependências) e depois o teste UNO real (mais lento,
precisa de 'soffice' no PATH). Para rodar só os rápidos (ex.: sem LibreOffice
disponível), use --sem-uno.

Uso:
    python completa/tests/run_all.py
    python completa/tests/run_all.py --sem-uno
"""
import subprocess
import sys
import os

RAIZ = os.path.dirname(os.path.abspath(__file__))

SMOKE_TESTS = [
    "smoke_test_unificacao.py",
    "smoke_test_extracao.py",
    "smoke_test_ganhos_rapidos.py",
    "smoke_test_ied.py",
]
TESTES_UNO = [
    "teste_uno_protocolos.py",
    "teste_paridade_import_export.py",
]


def rodar(script):
    caminho = os.path.join(RAIZ, script)
    print(f"\n=== {script} ===")
    resultado = subprocess.run([sys.executable, caminho])
    return resultado.returncode == 0


def main():
    sem_uno = "--sem-uno" in sys.argv
    scripts = list(SMOKE_TESTS) + ([] if sem_uno else list(TESTES_UNO))
    falharam = [s for s in scripts if not rodar(s)]
    print("\n" + "=" * 60)
    if falharam:
        print(f"FALHOU: {falharam}")
        sys.exit(1)
    print(f"OK: todos os {len(scripts)} testes passaram.")


if __name__ == "__main__":
    main()
