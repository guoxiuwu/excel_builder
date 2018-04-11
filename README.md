# excelbuilder

动态创建Excel 根据模板(DSL)文件，如下文件将产生一个Excel文件, 支持格式设置,分页传入数据等特性..

—————————————————————————————————————————————

excel {

    sheet("sheet1") {
        columns {
            column('AAAA', 'AAAA')
            skipCells(1)
            column('BBBB', 'BBBB')
        }
				
        row {
            cell([font: [bold: true, italic: false]], { dataContext.field1 })
            cell([font: [bold: true, italic: true]],  { dataContext.field2 })
            cell { dataContext.field3 }
            cell { dataContext.field4 }
            cell { dataContext.field5 == "field5" ? "--" : dataContext.field5 }
        }
				
        loop("registeredDataSource1") {
            row {
                cell { dataItem.id }
                cell { dataItem.id }
                cell { dataItem.id }
                cell { dataItem.id }
                cell { dataItem.id }
            }
        }
				
        skipRows(1)
				
        loop("registeredDataSource2") {
            row {
                cell { dataItem.id }
                cell { dataItem.id }
                cell { dataItem.id }
                cell { dataItem.id }
                cell { dataItem.id }
            }
        }
				
        skipRows(1)
				
        loop(dataContext.selfHoldList) {
            row {
                cell { dataItem.key }
                cell { dataItem.key }
                cell { dataItem.key }
                cell { dataItem.key }
                cell { dataItem.key }
            }
        }
    }
}
