# Automação de Follow-up 365

Projeto de portfólio em Python para transformar uma lista de pedidos pendentes em rascunhos personalizados no Microsoft Outlook. Cada fornecedor recebe apenas seus próprios itens, com tabela HTML e anexo correspondente.

## Destaques

- geração em lote de mensagens personalizadas;
- tabela HTML com pedidos pendentes;
- anexos individuais por fornecedor;
- preservação da assinatura configurada no Outlook;
- modo seguro: o script salva rascunhos e nunca envia mensagens automaticamente.

## Estrutura

```text
src/criar_rascunhos_outlook.py   lógica principal
examples/manifesto.exemplo.json  manifesto com dados fictícios
examples/pedidos_exemplo.csv     anexo fictício
```

## Como executar

Requisitos: Windows, Outlook desktop configurado e Python 3.10+.

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python src/criar_rascunhos_outlook.py --manifesto examples/manifesto.exemplo.json
```

Revise os rascunhos no Outlook antes de qualquer envio. Para usar dados próprios, copie o formato do manifesto de exemplo e mantenha arquivos reais fora do controle de versão.

## Privacidade

Esta edição pública contém apenas dados fictícios. Bases operacionais, auditorias, contatos, endereços corporativos e o manifesto original não fazem parte do repositório.

## Tecnologias

Python, JSON, HTML e automação COM do Microsoft Outlook (`pywin32`).
