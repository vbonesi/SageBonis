# -*- coding: utf-8 -*-
"""Roda os testes do SageBonis: o smoke test em memória (rápido, sem dependência) e
depois o teste UNO real (mais lento, precisa de 'soffice' no PATH). Para rodar só o
rápido (ex.: sem LibreOffice disponível), use --sem-uno.

Uso:
    python tests/run_all.py
    python tests/run_all.py --sem-uno
"""
import os
import subprocess
import sys

RAIZ = os.path.dirname(os.path.abspath(__file__))

SMOKE_TESTS = [
    "smoke_test_parser.py",
]
TESTES_UNO = [
    "teste_uno_roundtrip.py",
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
