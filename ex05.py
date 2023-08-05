
import xlsxwriter as xlsx
import os
caminho = 'C:\\Users\\FIC\\Documents\\izaias\\'
arquivo = 'exercicio5.xlsx'


# -------------------
workbook = xlsx.Workbook(caminho + arquivo)

sheetdados = workbook.add_worksheet('relatorio')

dados = [
  ['Produto','Qtd','Preço','Subtotal'],
  ['calça',5,50,],
  ['camisa',4,30,],
  ['Sapato',5,40,],
  ['saia',8,20,],
  ['Blusa',6,10,],
  ['shorts',8,25,]]

# fomato data
formatodata = workbook.add_format({'bg_color':'#C5D4C9',
                                'font_color':'black',
                                'align':'center',
                                'bold':True,
                                'num_format':'dd/mm/yyyy'})

# formato hora
formatohora = workbook.add_format({'bg_color':'#C5D4C9',
                                'font_color':'black',
                                'align':'center',
                                'bold':True,
                                'num_format':'h:mm'})



# titulos da parte azul
titles = workbook.add_format({'bg_color':'#332FD3',
                                'font_color':'white',
                                'bold':True,
                                'font_name':'Arial'
                                   })
# coluna de quantidades
qtde= workbook.add_format({'bg_color':'#332FD3',
                                'font_color':'white',
                                'bold':True,
                                'font_name':'Arial'
                                    })
# cor do cabeçalho e total 
formatacaocabecalho = workbook.add_format({'bg_color':'#C5D4C9',
                                'font_color':'black',
                                'font_size': 12,
                                'align':'center',
                                'bold':True,
                                'font_name':'Arial'
                                          })
# formatação dos valores preços e resultados 
cortotais = workbook.add_format({
                                'num_format':'R$ #,##0',
                                'fg_color':'#55FBEF',
                                'font_color':'#100A0D',
                                'bold':True,
                                'font_name':'Arial'
                                })

sheetdados.merge_range('A1:D1', 'planilha de solicitações',formatacaocabecalho)

sheetdados.write('A3','Data:')
sheetdados.write_formula('B3','=Today()',formatodata)


sheetdados.write('A4','Hora:')
sheetdados.write_formula('B4','=Now()',formatohora)
sheetdados.write('C13','TOTAL:')
sheetdados.write_formula('D13','=sum(D7:D12)',cortotais)

poslinha = 5
poscoluna = 0 

for linha, listas in enumerate(dados):
  if linha == 0 :
    sheetdados.write_row(linha+ poslinha, poscoluna,listas)
  else:
    sheetdados.write(linha+ poslinha, poscoluna,listas[0])#produto
    sheetdados.write(linha+ poslinha, poscoluna + 1,listas[1])#qtde
    sheetdados.write(linha+ poslinha, poscoluna + 2,listas[2],cortotais)#preço unitario
    sheetdados.write_formula(linha + poslinha , poscoluna + 3, '=B' + str(linha + poslinha+1) + '*c' + str(linha + poslinha+1),cortotais)



sheetdados.set_column,('A:D',35.5)
workbook.close()
os.startfile(caminho + arquivo)












