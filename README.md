# Prontuário: versões publicadas

Este repositório guarda **somente as versões publicadas** do Prontuário: o arquivo
`prontuario.zip` e a soma `SHA256SUMS.txt` de cada versão, em **Releases**. Nenhum dado de
paciente passa por aqui, e o código-fonte fica em um repositório privado.

- **Baixar a versão mais recente:** [Releases → Latest](../../releases/latest).
- **Atualizar:** a partir da versão 0.2.6, o `Prontuario.bat` e o `Agenda.bat` conferem esta página
  ao abrir e oferecem a versão nova ("Atualizar agora? [S/n]"), conferindo a soma SHA-256 antes de
  instalar.

As versões são publicadas automaticamente pelo processo de publicação do repositório principal.

## Ferramentas

- [`ferramentas/Inspecionar-HiDoctor.cmd`](ferramentas/Inspecionar-HiDoctor.cmd): gera, no próprio
  computador, um relatório só com a **estrutura** do banco do HiDoctor (programas, nomes de tabelas e
  colunas, contagens), para preparar a importação. Não lê nenhum registro de paciente e não envia
  nada. Para baixar: abra o link e clique no botão **Download raw file** (seta para baixo).

