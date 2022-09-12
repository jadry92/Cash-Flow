
interface rowInterface {
  description: string;
  date: Date;
  frequency: string;
  value: number;
  variance: number;
}

interface rowDataBase {
  date: Date;
  value: number;
  description: string;
}

interface savings {
  date: Date;
  value: number;
}

interface dataBase {
  [index: string]: rowDataBase[]
}


function generateDataBase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const gr = new GraphCashFlow(ss);
  gr.readDataFromInterface()
  gr.cleanDataBase()
  gr.generateCashFlowData()
  gr.render()
}

class GraphCashFlow {
  private sDataBase: GoogleAppsScript.Spreadsheet.Sheet;
  private sInteraction: GoogleAppsScript.Spreadsheet.Sheet;
  // data
  initialSavings: savings;
  dataInterface: [rowInterface];
  fullData: dataBase;
  fullDataRaw: rowDataBase[];
  // const
  numberOfMonths: number;
  initialDate: Date;
  endDate: Date;


  constructor(ss) {
    // spread sheet with all the data
    this.sDataBase = ss.getSheets()[1]
    // spread with the graph and the interaction with the user
    this.sInteraction = ss.getSheets()[0]
  }


  readDataFromInterface() {
    let flag = true
    let row = 2
    // get seed data
    while (flag) {
      const dataCells = this.sInteraction.getRange(`A${row}:E${row}`)

      if (dataCells.getValues()[0][0] != '') {
        const data = {
          description: dataCells.getValues()[0][0],
          date: new Date(dataCells.getValues()[0][1]),
          frequency: dataCells.getValues()[0][2],
          value: parseInt(dataCells.getValues()[0][3]),
          variance: parseFloat(dataCells.getValues()[0][4])
        }
        row === 2 ? this.dataInterface = [data] : this.dataInterface.push(data)
        row++
      } else {
        flag = false
      }
    }
    // get constants
    this.numberOfMonths = this.sInteraction.getRange('I3').getValue()
    if (!this.numberOfMonths) {
      Logger.log("Number Of Months is not define")
      throw new Error("Number Of Months is not define")
    }
    this.initialDate = new Date(this.sInteraction.getRange('I2').getValue())
    if (!this.initialDate) {
      Logger.log("Initial Date is not define")
      throw new Error("Initial Date is not define")
    }
    this.endDate = new Date(this.initialDate.getFullYear(), this.initialDate.getMonth() + this.numberOfMonths - 1, this.initialDate.getDate()) 
    // get initial savings
    const savingsCells = this.sInteraction.getRange('I1:J1')
    this.initialSavings = {
      date: new Date(savingsCells.getValues()[0][0]),
      value: parseInt(savingsCells.getValues()[0][1])
    }
    if (!this.initialSavings){
      Logger.log("savings is not define")
      throw new Error("savings is not define")
    }
  }

  generateCashFlowData() {
    //generate
    for (const seed in this.dataInterface) {
      const rows = this.frequencySwitch(this.dataInterface[seed])
      if (this.fullDataRaw !== undefined) {
        this.fullDataRaw = this.fullDataRaw.concat(rows)
      } else {
        this.fullDataRaw = rows
      }
    }
    this.fullDataRaw.sort(function (a: rowDataBase, b: rowDataBase) {
      return a.date.getTime() - b.date.getTime()
    })
    //sort

    for (let monthNum = 0; monthNum < this.numberOfMonths; monthNum++) {
      const date = new Date(this.initialDate.getFullYear(), this.initialDate.getMonth() + monthNum, this.initialDate.getDate())
      const monthIndex = this.indexDate(date)
      if (monthNum === 0) {
        this.fullData = {
          [monthIndex]: [{}] as rowDataBase[]
        }
      } else {
        this.fullData[monthIndex] = [{}] as rowDataBase[]
      }
    }

    for (const row in this.fullDataRaw) {
      if (this.fullDataRaw[row].date < this.endDate) {
        const monthIndex = this.indexDate(this.fullDataRaw[row].date)
        this.fullData[monthIndex].push(this.fullDataRaw[row])
      }
    }
    this.calculateSavings()
    /*
        const date = new Date(this.initialDate.getFullYear(), this.initialDate.getMonth() + monthNum, this.initialDate.getDate())
        const options: { month: 'short' } = { month: 'short' }
        const monthName = new Intl.DateTimeFormat('en-US', options).format(date)
    */
  }

  private dateToString(date: Date): string {
    const day = date.getDate()
    const month = date.getMonth() + 1
    const year = date.getFullYear()
    return `${day}/${month}/${year}`
  }

  private indexDate(date: Date): string {
    const year = date.getFullYear()
    const month = date.getMonth() + 1
    return `${month}/${year}`
  }

  private frequencySwitch(seed: rowInterface): [rowDataBase] {
    let data: [rowDataBase] 
    switch (seed.frequency) {
      case 'monthly':
        for (let month = 0; month < this.numberOfMonths; month++) {
          let date = new Date(seed.date.getFullYear(), seed.date.getMonth() + month, seed.date.getDate())
          const row = {
            date: date,
            value: seed.value,
            description: seed.description
          }
          month === 0 ? data = [row] : data.push(row)
        }
        break;
      case 'fortnightly':
        const fortnights = Math.floor(Math.floor((365.0 / 12.0) * this.numberOfMonths) / 14.0)
        for (let fortnight = 0; fortnight < fortnights; fortnight++) {
          let date = new Date(seed.date.getFullYear(), seed.date.getMonth(), seed.date.getDate() + (14 * fortnight))
          const row = {
            date: date,
            value: seed.value,
            description: seed.description
          }
          fortnight === 0 ? data = [row] : data.push(row)
        }
        break;
      case 'weekly':
        const weeks = Math.floor(Math.floor((365.0 / 12.0) * this.numberOfMonths) / 7.0)
        for (let week = 0; week < weeks; week++) {
          let date = new Date(seed.date.getFullYear(), seed.date.getMonth(), seed.date.getDate() + (7 * week))
          const row = {
            date: date,
            value: seed.value,
            description: seed.description
          }
          week === 0 ? data = [row] : data.push(row)
        }
        break;
      case 'once':
        data = [{
          date: seed.date,
          value: seed.value,
          description: seed.description
        }]
        break;
      default:
        data = [{
          date: new Date(),
          value: 0,
          description: 'null'
        }]
    }
    return data
  }

  writeCashFlowData() {

  }

  private calculateSavings() {
    let month = 0

    let oldKey: string
    for (const key in this.fullData) {
      let savingsValue = 0
      let date: Date
      if (month === 0) {
        date = this.initialSavings.date
        savingsValue = this.initialSavings.value
      } else {
        const rows = this.fullData[oldKey]
        for (const row in rows) {
          if (rows[row].value != undefined) {
            savingsValue += rows[row].value
          }
        }
        date = new Date(this.initialSavings.date.getFullYear(), this.initialSavings.date.getMonth() + month, 1)
      }

      this.fullData[key][0] = {
        date: date,
        value: savingsValue,
        description: 'Savings'
      }
      month++
      oldKey = key
    }

    // last savings
    let endDateSavings = new Date(this.endDate.getFullYear(), this.endDate.getMonth() + 1, 1)

    let savingsValue = 0

    const rows = this.fullData[oldKey]
    for (const row in rows) {
      if (rows[row].value != undefined) {
        savingsValue += rows[row].value
      }
    }

    this.fullData[this.indexDate(endDateSavings)] = [{
      date: endDateSavings,
      value: savingsValue,
      description: 'Savings'
    }]
  }

  cleanDataBase() {
    const range = this.sDataBase.getRange('A2:D')
    range.clear()
  }

  render() {
    // This function render the graph from the database.
    let idMonth = 2
    let idRow = 2
    for (const key in this.fullData) {
      const rows = this.fullData[key]
      for (const row in rows) {
        const cellRows = this.sDataBase.getRange(`B${idRow}:D${idRow}`)
        const values = [
          [this.dateToString(rows[row].date), rows[row].value.toString(), rows[row].description]
        ]
        cellRows.setValues(values)
        idRow++
      }
      const cellMonths = this.sDataBase.getRange(`A${idMonth}:A${idRow - 1}`)
      cellMonths.merge()
      cellMonths.setValue(key)
      idMonth = idRow
    }
  }

}
