#注意事项：1.execl中都是有效表格数据。2.行列都顶头写
#导入操作ececl和word文档的库
from docx import Document
from docx.shared import Inches
from docx.enum.style import WD_STYLE_TYPE
import xlrd

#打开execl
execl = xlrd.open_workbook('demoececl.xlsx')

#打开word文档
document = Document()

#计算execl中sheet的个数
sheetCount = len(execl.sheets())
#print(sheetCount)
for index in range(0,sheetCount):
    table = execl.sheets()[index]
    #获取表格的行数和列数
    tableRows = table.nrows
    tableCols = table.ncols
    
    #保存合并单元格的信息
    mergedInfo = []
    mergedDict = {'rs':0,'re':0,'cs':0,'ce':0}
    #获取表格中合并的单元的信息(起始行，结束行，起始列，结束列)
    for crange in table.merged_cells:
        temp = mergedDict.copy()
        temp['rs'], temp['re'], temp['cs'], temp['ce'] = crange
        #不能直接 mergedInfo.append(mergedDict)
        #因为因为字典d 是一个object ,而mergedInfo.append(mergedDict)并没有真正的将该字典在内存中再次创建。只是指向了相同的object。
        #这也是python 提高性能，优化内存的考虑。
        mergedInfo.append(temp)
        
    i = 0
    writeRe = 0
    while i < tableRows:
        isExist = False
        if len(mergedInfo) > 0:
            for j in range(0,len(mergedInfo)):
                if i == mergedInfo[j]['rs']:
                    mergeRow = mergedInfo[j]['re'] - mergedInfo[j]['rs']
                    isExist = True
                    break
        if isExist == False:
            mergeRow = 1
        
        writeRe = i + mergeRow 
        wordTable = document.add_table(rows=mergeRow, cols=tableCols - 1,style="Table Grid")
        document.add_paragraph()
  
        #每个表格开始写的时候行索引都为0
        rowIndex = 0
        #向表格写数据 
        for row in range(i,writeRe):
            #每行开始写数据时，列索引都为0
            colIndex = 0
            for col in range(1,tableCols):
                data = table.cell(row,col)
                if data.value is not None:
                    try:
                        writeTableData = str("")
                        if isinstance(data.value,float):
                            if  round(data.value) == data.value:
                                writeTableData = str(round(data.value))
                            else:
                                writeTableData = str(data.value)
                        else:
                            writeTableData = data.value
                        wordTable.cell(rowIndex,colIndex).text = str(writeTableData)
                        colIndex = colIndex + 1
                    except IndexError:
                        pass
            rowIndex = rowIndex + 1
        i = i + mergeRow

document.save("test1.docx")