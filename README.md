Renomeador de PDFs

Este código foi criado para ajudar a minha namorada a resolver um problema específico:

Ela tinha mais de 5 mil PDFs com nomes genéricos (ex: ARQUIVO1.pdf) e precisava renomeá-los para um formato específico, utilizando dados contidos dentro desses PDFs. Para isso, desenvolvi este "bot", que extrai as informações dos PDFs e as transforma em tabelas no Excel. Com essas tabelas, consegui obter os dados necessários para renomear os arquivos de acordo com os padrões e condições que ela estabeleceu.

Sim, o código pode parecer uma bagunça e foi feito para resolver esse caso específico, mas acredito que ele possa servir como um guia para quem deseja criar algo similar.

O formato final dos arquivos, conforme as informações extraídas dos PDFs, segue este padrão:
{tema}_{peca}_{ref}_{ultima_palavra}.pdf = ACADEMIA INOVAR_POLO ESPECIAL_12773

Além disso, havia uma condição extra: se o nome gerado fosse duplicado, o código adicionava um contador no final, resultando em algo como:
{tema}_{peca}_{ref}_{ultima_palavra} ({counter}).pdf = ACADEMIA INOVAR_POLO ESPECIAL_12773 (2)

Variáveis principais:
input_folder: pasta onde os arquivos originais são lidos.
output_folder: pasta onde os arquivos renomeados são salvos.
