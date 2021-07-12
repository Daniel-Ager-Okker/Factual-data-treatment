import xlrd
import xlsxwriter
def dash_type(c): #это функция, которая отвечает за тип кривой изменения эксплуатационного показателя
  if 'H' in c:
    return 'dash'
  else:
    return 'solid'
print('Программа предназначена для отстраивания графиков по данным, выгруженным из CETStudio')
print('Введит полное имя входного файла')
print('Данные должны находиться на 1-ом листе')
file_name=input()
File_In=xlrd.open_workbook(file_name)
Sheet_In=File_In.sheet_by_index(0)
n_rows=Sheet_In.nrows
n_cols=Sheet_In.ncols
File_Out=xlsxwriter.Workbook('Out.xlsx')
Sheet_Out=File_Out.add_worksheet()
Sheet_Out.write(0,0,Sheet_In.cell_value(0,0)) #скопировали наименование столбца "время"
for j in range(1,n_rows):
    Sheet_Out.write(j,0,Sheet_In.cell_value(j,0)) #скопировали сам столбец с датами (значения)
for i in range(1,n_cols,2):
  Sheet_Out.write(0,i,Sheet_In.cell_value(0,i)) #скопировали наименование столбца "расчетный показатель"
  Sheet_Out.write(0,i+1,Sheet_In.cell_value(0,i+1)) #скопировали наименование столбца "фактичекий показатель"
  for j in range(1,n_rows):
    Sheet_Out.write(j,i,Sheet_In.cell_value(j,i)) #скопировали сам столбец с значениями расчетных эксплуатационных показателей
    Sheet_Out.write(j,i+1,Sheet_In.cell_value(j,i+1)) #скопировали сам столбец с значениями фактических эксплуатационных показателей 

pos=1
for i in range(1,n_cols,2):
  chart=File_Out.add_chart({'type':'line'})
  #добавялем на график расчетные данные
  chart.add_series({
      'categories':['Sheet1',1,0,n_rows,0],
      'values':['Sheet1',1,i,n_rows,i],
      'name':'Забойное давление',
      'line':{'dash_type':'solid','color':'b686cb','width':1.5}
      #'marker':{'type':'circle', 'size':4}
  })
  #дбавляем на график фактические данные
  chart.add_series({
      'categories':['Sheet1',1,0,n_rows,0],
      'values':['Sheet1',1,i+1,n_rows,i+1],
      'name':'Забойное давление история',
      'line':{'dash_type':'long_dash','color':'black','width':1.5}
      #'marker':{'type':'circle', 'size':4}
  })
  #chart.set_plotarea({'pattern':{'pattern':'large_grid','fg_color':'#dadada','bg_color':'white'}})
  chart.set_x_axis({'num_font':{'rotation':0}}) #сделали так, чтобы подписи по оси Х были горизонтальными
  chart.set_y_axis({'name':'забойное давление, бар','bold':False})
  chart.set_legend({'none':False,'position':'bottom'})
  Sheet_Out.insert_chart(5,n_cols+pos,chart)
  pos+=8
File_Out.close()