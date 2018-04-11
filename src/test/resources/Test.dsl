excel {
    sheet("sheet1") {

        columns {
            column('AAAA', 'AAAA')
            skipCells(1)
            column('BBBB', 'BBBB')
            column('CCCC', 'CCCC')
            column('DDDD', 'DDDD')
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