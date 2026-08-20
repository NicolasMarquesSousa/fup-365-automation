"""Cria rascunhos de follow-up no Outlook a partir de um manifesto sanitizado."""

from __future__ import annotations

import argparse
import html
import json
from pathlib import Path

import win32com.client


HEADERS = ["DESCRIÇÃO", "SALDO", "PEDIDO", "EMISSÃO", "FORNECEDOR", "QTD.", "VALOR", "TOTAL"]


def tabela_html(linhas: list[list[object]]) -> str:
    cabecalho = "".join(f"<th>{html.escape(item)}</th>" for item in HEADERS)
    corpo = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(valor))}</td>" for valor in linha) + "</tr>"
        for linha in linhas
    )
    estilo = "border:1px solid #777;padding:5px;text-align:left"
    return f'<table style="border-collapse:collapse;font-family:Arial;font-size:11pt"><tr>{cabecalho}</tr>{corpo}</table>'.replace("<th>", f'<th style="{estilo}">').replace("<td>", f'<td style="{estilo}">')


def corpo_html(item: dict, assinatura: str) -> str:
    return f"""<div style="font-family:Arial;font-size:11pt">
    <p>Olá, equipe do fornecedor.</p>
    <p>Os pedidos abaixo estão sem atualização. Pedimos a revisão do status no portal.</p>
    {tabela_html(item['pedidos'])}
    <p>Este rascunho deve ser revisado antes do envio.</p>
    </div>{assinatura}"""


def criar_rascunhos(manifesto: Path, cc_padrao: str = "") -> int:
    itens = json.loads(manifesto.read_text(encoding="utf-8"))
    outlook = win32com.client.Dispatch("Outlook.Application")
    criados = 0
    for item in itens:
        mensagem = outlook.CreateItem(0)
        mensagem.Display(False)
        assinatura = mensagem.HTMLBody
        mensagem.To = item["para"]
        mensagem.CC = cc_padrao
        mensagem.Subject = item["assunto"]
        mensagem.HTMLBody = corpo_html(item, assinatura)
        anexo = (manifesto.parent / item["anexo"]).resolve()
        if not anexo.is_file():
            raise FileNotFoundError(f"Anexo não encontrado: {anexo}")
        mensagem.Attachments.Add(str(anexo))
        mensagem.Save()  # Segurança: cria rascunho; não envia automaticamente.
        mensagem.Close(0)
        criados += 1
        print(f"[RASCUNHO] {item['fornecedor']} ({len(item['pedidos'])} pedido(s))")
    return criados


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifesto", type=Path, default=Path("examples/manifesto.exemplo.json"))
    parser.add_argument("--cc", default="", help="CC opcional; não mantenha endereços reais no código")
    args = parser.parse_args()
    total = criar_rascunhos(args.manifesto, args.cc)
    print(f"Concluído: {total} rascunho(s). Nenhum e-mail foi enviado.")
