# -*- coding: utf-8 -*-
"""A régua dos testes de round-trip: a gramática do .dat como o CARREGADOR DO SAGE a lê.

Medida no SAGE 31 (container Docker, `AtualizaBD fria fonte`, ver tests/docker/):
  * um registro começa numa linha que é SÓ o nome da entidade e vai até o próximo nome;
  * linha em branco e linha que começa com ';' são ignoradas -- inclusive ';ENTIDADE' e
    ';#include' -- e NÃO encerram o registro;
  * 'chave = valor' continua o registro corrente (mesmo depois de comentários);
  * atributo com valor vazio equivale a atributo ausente;
  * '#include' encerra o registro.
Dois textos que dão os mesmos registros são a mesma base para o SAGE, ainda que o arquivo
mude de formato (ordem de atributos, comentários, espaçamento, linhas em branco).
"""
import re


def registros_do_carregador(texto):
    """[(ENTIDADE, tuple(atributos ordenados e não vazios))] + ('#include', alvo).
    Atributo repetido: vale o último (é o que a macro faz; o carregador do SAGE o rejeita)."""
    saida, atual = [], None
    for linha in texto.replace("\r", "").split("\n"):
        s = linha.strip()
        if not s or s.startswith(";"):
            continue
        m = re.match(r"#\s*include\s+(.*)", s, re.I)
        if m:
            saida.append(("#include", m.group(1).strip().replace("\\", "/")))
            atual = None
            continue
        if re.fullmatch(r"[A-Za-z0-9_]+", s):
            atual = [s.upper(), {}]
            saida.append(atual)
            continue
        if "=" in s and atual is not None:
            k, v = s.split("=", 1)
            if v.strip():
                atual[1][k.strip()] = v.strip()
    return [(r[0], tuple(sorted(r[1].items()))) if isinstance(r, list) else r for r in saida]


def ler_latin1(caminho):
    with open(caminho, "rb") as f:
        return f.read().decode("latin-1")
