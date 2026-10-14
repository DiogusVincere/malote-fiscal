"""Malote Fiscal: pergunta aos clientes, pelo Outlook, se tomaram serviços no mês, guarda as notas que
eles mandam na pasta de cada um e ajuda a responder.

- app.py: a tela (Tkinter). `python malote.pyw` abre; `--demo` abre com o Outlook simulado.
- outlook.py: tudo o que conversa com o Outlook (COM) e a lógica pura de leitura das respostas.
- falso.py: o Outlook simulado e o escritório fictício da demonstração e dos testes.
- demo.py: o fluxo inteiro sem tela (`python -m malote.demo`), para ver e para a integração contínua.
- planilha.py: a planilha de clientes (quem recebe e de quem é cada e-mail).
- cadastro.py: o cadastro próprio de contatos, que a planilha e os grupos atualizam e a pessoa corrige.
- clientes.py: índice das pastas de clientes, quem é o remetente e onde salvar.
- competencia.py: mês de referência, assunto do e-mail, mês citado na resposta.
- devolucao.py: avisos de "não entregue".
- anexos.py: CNPJ dentro do anexo, para e-mail que serve a várias empresas.
- sugestao.py: sugestão de resposta com aprendizado supervisionado, sem IA de fora.
- layouts.py: texto da pergunta e respostas prontas, editáveis.
- registro.py: o que a ferramenta aprendeu e o que já salvou.
- controle.py: as planilhas de controle do mês.
- tema.py: paleta e estilos ttk.
"""
VERSAO = "1.0.0"
