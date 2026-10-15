# Cartazes de qualificação de doutorado

Automação do cartaz enviado, com o mesmo fluxo de planilha → documento → imagem
da automação de mestrado. O modelo mantém a arte original e usa caixas de texto
ancoradas na página, com posição e tamanho fixos. Títulos ou nomes maiores não
empurram os demais campos nem sobrepõem o rodapé.

O script mede as linhas e reduz somente a fonte da caixa que precisa de espaço.
O ajuste abrange título, autoria, orientador, banca, data, hora e local. Na
exportação, confere o texto visível de cada caixa no PDF e, se necessário, reduz
a fonte novamente antes de gerar a imagem. O tamanho mínimo é 8 pontos; textos
que ainda não caibam causam erro, sem entregar um cartaz cortado.

## Uso no Windows

1. Instale Python 3.10 ou superior e Microsoft Word ou LibreOffice.
2. Edite `dados_qualificacao_doutorado.xlsx`, aba **Dados**: uma linha por aluno.
   A linha de Adrian é um exemplo; substitua-a ou exclua-a se não quiser gerar seu cartaz.
3. Feche a planilha e dê dois cliques em `gerar_cartazes.bat`.
4. Abra `cartazes_gerados`: cada aluno terá seu DOCX, PDF e PNG em 300 dpi.

O arquivo BAT instala as bibliotecas em uma pasta `.venv` e exibe eventuais erros.
Uma nova execução protege as pastas existentes: renomeie a pasta de saída ou use
`--sobrescrever` conscientemente pela linha de comando.

## Uso pela linha de comando (Linux, macOS e Windows)

Execute estes comandos nesta pasta:

```sh
python -m venv .venv
# Linux/macOS:
source .venv/bin/activate
# Windows, em vez do comando acima:
# .venv\Scripts\activate
python -m pip install -r requirements.txt
python run.py
```

No Linux é necessário LibreOffice para gerar PDF/PNG. No Debian/Ubuntu, instale
`libreoffice-writer` e `fonts-crosextra-carlito` pelo gerenciador de pacotes.
Carlito é uma alternativa com métricas compatíveis com Calibri. A renderização
pode variar entre Word e LibreOffice; confira o cartaz antes de divulgar.
No Windows/macOS o script prefere LibreOffice quando disponível; caso contrário,
usa Microsoft Word por `docx2pdf`.

```sh
# Apenas documentos editáveis, sem exigir Word ou LibreOffice:
python run.py --formato docx
# Planilha e saída diferentes:
python run.py --planilha meus_dados.xlsx --saida nova_saida
# Selecionar o conversor:
python run.py --conversor word
python run.py --conversor libreoffice --soffice "/caminho/para/soffice"
# Atualizar arquivos previamente gerados:
python run.py --sobrescrever
```

Também é possível definir `SOFFICE_PATH` com o caminho do LibreOffice.
Os caminhos padrão são relativos à pasta do script, independentemente de onde
ele for chamado. Nenhum serviço ou credencial é necessário.

## Preenchimento da planilha

Preserve os cabeçalhos. São obrigatórios **Aluno**, **Título**, **Orientador**,
**Presidente**, **Membro Interno**, **Membro Externo 1**, **Data**, **Hora** e **Sala**.
Os demais membros externos e os dois suplentes são opcionais; posições vazias
não exibem marcadores, e a caixa da banca mantém as mesmas dimensões.
Universidades ficam em colunas separadas.
Não escreva a função (Presidente, Suplente etc.) no nome: o script a adiciona.

Digite os nomes dos professores com a titulação desejada, como `Profª. Drª. Nome`.
O script preserva a grafia e não tenta deduzir gênero ou alterar nomes.
A planilha de exemplo transcreve os dados do documento recebido, inclusive
titulações e grafia: revise-os antes de divulgar.

- **Data**: `31/03/2026`, `2026-03-31` ou uma célula de data do Excel.
- **Hora**: `14:00`, `14h00` ou uma célula de hora do Excel.
- **Sala**: descrição completa, como `Sala NEB-02 / NEB-UFPA`.
- **Fonte Título**: tamanho máximo desejado, entre 8 e 36 pontos. Se vazia,
  o máximo é 21. O script reduz o tamanho quando necessário para caber na caixa,
  inclusive quando esta coluna está preenchida.

Use valores diretos, sem fórmulas. Alunos duplicados, campos obrigatórios vazios
e datas inválidas impedem a execução antes de gerar documentos.
O modo somente DOCX ajusta as linhas por métricas de fonte, com uma margem de
segurança. A conferência do resultado efetivamente renderizado requer a geração
de PDF/PNG com Word ou LibreOffice. Se a fonte mínima não bastar, reduza o texto
na planilha. O script nunca aumenta a caixa para acomodar conteúdo.

Ao atualizar uma instalação antiga, substitua também
`modelo_qualificacao_doutorado.docx` e inclua `layout.py`. O script novo exige o
modelo com caixas fixas. Preserve sua planilha preenchida; seus dados continuam
compatíveis com esta versão.

O script retorna código diferente de zero quando há falhas e informa quais
linhas falharam. Os arquivos de cada cartaz são preparados em uma pasta temporária
e só são entregues após a geração completa desse cartaz. Linhas anteriores que
passaram permanecem disponíveis se uma linha posterior falhar.

## Validação

```sh
python -m unittest discover -s tests -v
```

Os testes verificam preenchimento, caracteres especiais, preservação da arte,
datas/horários, proteção de arquivos e rejeição de dados inválidos. Também
verificam que textos maiores reduzem a fonte sem mudar a geometria das caixas.
Se LibreOffice estiver disponível (no PATH ou em `SOFFICE_PATH`), executam a
conversão de exemplos curtos e longos, conferindo todo o texto e as posições dos
retângulos no PDF. Esse teste é marcado como ignorado quando não há LibreOffice.
O exemplo também deve ser executado com `python run.py --saida /caminho/de/teste`
para validar o conversor instalado e conferir visualmente o resultado.
