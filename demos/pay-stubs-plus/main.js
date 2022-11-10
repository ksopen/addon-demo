import './app-jssdk/index.js'

const jssdkApi = window.jssdk.api
await jssdkApi.ready()
const Application = jssdkApi.Application

const BASE_INFO = ["工号", "姓名", "职位"]
const PAY_INFO = [
  {
    header: '应领工资',
    items: ["基本工资", "全勤奖", "岗位奖金", "职务津贴", "加班工资", "业务佣金", "合计金额"]
  },
  {
    header: '应扣工资',
    items: ["社保", "税金", "事假", "病假", "旷工", "代扣", "合计金额"]
  }
]
const ONE_LINE_HEADER = ['工号', '姓名', '职位', '岗位工资', '绩效工资', '教护龄津贴', '边远地区津贴', '乡村教师生活补贴', '代扣公积金', '实发工资']
let totalLen = 18
let headerRow = 1

// 创建模板
async function createTemplate() {
  // 新建工作表
  const sheets = Application.ActiveWorkbook.Sheets
  // 添加工作表
  await sheets.Add(null, null, 1, Application.Enum.XlSheetType.xlWorksheet, '工资条模板')
  // 生成模板基本信息区域
  let firstPart = Application.Range('A1:A2')
  for (let i = 0; i < BASE_INFO.length; i++) {
    await firstPart.Merge()
    firstPart.Value = BASE_INFO[i]
    firstPart = await firstPart.Offset(0, 1)
  }

  // 获取表格头部的两行
  let firstRowCell = await firstPart.Item(1)
  let secondRowCell = await firstPart.Item(2)

  // 根据数据进行单元格数据填充
  for (let step = 0; step < PAY_INFO.length; step++) {
    let firstRowEndCell = firstRowCell
    let i = 0
    const len = PAY_INFO[step].items.length
    for (; i < len; i++) {
      secondRowCell.Value = PAY_INFO[step].items[i]
      secondRowCell = await secondRowCell.Offset(0, 1)
    }

    // 将选择区域右移
    firstRowEndCell = await firstRowCell.Offset(0, i - 1)
    // 拼接选择区域的字符串，如：$A$1:$B$2
    const rangeStr = `${await firstRowCell.Address()}:${await firstRowEndCell.Address()}`

    const basePayHeaderRange = Application.Range(rangeStr)
    await basePayHeaderRange.Merge()
    basePayHeaderRange.Value = PAY_INFO[step].header
    firstRowCell = firstRowEndCell.Offset(0, 1)
  }

  const endRangeStr = `${await firstRowCell.Address()}:${await secondRowCell.Address()}`
  const endRange = Application.Range(endRangeStr)
  await endRange.Merge()
  endRange.Value = '实发工资'

  let item = Application.Range('A4')
  for (let i = 0; i < ONE_LINE_HEADER.length; i++) {
    item.Value = ONE_LINE_HEADER[i]
    item.WrapText = true
    item = await item.Offset(0, 1)
  }
}
// 创建工资条的头部
async function createHeader(startRow, { baseInfo, payInfo, endInfo }) {
  if (headerRow === 1) {
    let item = Application.Range(`A${startRow}`)
    for (let i = 0; i < totalLen; i++) {
      item.Value = baseInfo[i]
      item.WrapText = true
      item = await item.Offset(0, 1)
    }
  }
  if (headerRow === 2) {
    // 生成模板基本信息区域
    let firstPart = Application.Range(`A${startRow}:A${startRow + 1}`)
    for (let i = 0; i < baseInfo.length; i++) {
      await firstPart.Merge()
      firstPart.Value = baseInfo[i]
      firstPart.WrapText = true
      firstPart = await firstPart.Offset(0, 1)
    }

    // 应领工资区域
    let firstRowCell = await firstPart.Item(1)
    let secondRowCell = await firstPart.Item(2)

    for (let step = 0; step < payInfo.length; step++) {
      let firstRowEndCell = firstRowCell
      let i = 0
      const len = payInfo[step].items.length
      for (; i < len; i++) {
        secondRowCell.Value = payInfo[step].items[i]
        secondRowCell.WrapText = true
        secondRowCell = await secondRowCell.Offset(0, 1)
      }

      firstRowEndCell = await firstRowCell.Offset(0, i - 1)
      const rangeStr = `${await firstRowCell.Address()}:${await firstRowEndCell.Address()}`
      const basePayHeaderRange = Application.Range(rangeStr)
      await basePayHeaderRange.Merge()
      basePayHeaderRange.Value = payInfo[step].header
      firstRowCell = firstRowEndCell.Offset(0, 1)
    }

    const endRangeStr = `${await firstRowCell.Address()}:${await secondRowCell.Address()}`
    let endRange = Application.Range(endRangeStr)
    for (let i = 0; i < endInfo.length; i++) {
      await endRange.Merge()
      endRange.Value = endInfo[i]
      endRange.WrapText = true
      endRange = await endRange.Offset(0, 1)
    }
  }

}
// 获取工资条头部数据
async function getHeaderInfo() {
  const selection = await Application.Selection
  const colums = await selection.ColumnEnd - await selection.Column + 1
  const rows = await selection.RowEnd - await selection.Row + 1
  headerRow = rows
  totalLen = colums
  const baseInfo = []
  const payInfo = []
  const endInfo = []
  const cells = await selection.Cells
  let part = 1
  if (headerRow === 1) {
    for (let i = 1; i <= colums; i++) {
      let item = await cells.Item(1, i)
      baseInfo.push(await item.Text)
    }
  } else {
    console.log(part);
    for (let i = 1; i <= colums; i++) {
      console.log(i);
      let item = await cells.Item(1, i)
      if (await item.MergeCells) {
        const MergeArea = await item.MergeArea
        const colums = await MergeArea.ColumnEnd - await MergeArea.Column + 1
        console.log('colums', colums);
        if (colums > 1) {
          part = 2
          payInfo.push({
            header: await item.Text,
            items: []
          })
          const len = payInfo.length
          for (let j = 0; j < colums; j++) {
            let nextRowCell = await cells.Item(2, j + i)
            payInfo[len - 1].items.push(await nextRowCell.Text)
          }
          i += (colums - 1)
        } else {
          if (part === 1) {
            baseInfo.push(await item.Text)
          } else {
            endInfo.push(await item.Text)
          }
        }
      }
    }
  }

  return {
    baseInfo,
    payInfo,
    endInfo
  }
}

// 获取工资数据
async function getData() {
  const data = []
  let range = Application.Range(`A${headerRow + 1}:${String.fromCharCode(65 + totalLen)}3`)
  let step = 0
  while (true) {
    const cells = await range.Cells
    const text = await cells.Item(1, 1).Text
    console.log(text)
    if (!text) {
      break
    }
    data.push([])
    for (let i = 1; i <= totalLen; i++) {
      let item = await cells.Item(1, i)
      data[step].push(await item.Text)
    }
    console.log(data[step])
    range = await range.Offset(1, 0)
    step++
  }

  return data
}
// 设置表格边框
async function setBorder(rowStart) {
  const endCellStr = String.fromCharCode(65 + totalLen - 1)
  const totalArea = Application.Range(`A${rowStart}:${endCellStr}${rowStart + headerRow}`)
  await totalArea.Select()
  const borders = totalArea.Borders
  // 设置区域中间的边框
  const borderInner = await borders.Item(Application.Enum.XlBordersIndex.xlInside)
  borderInner.Color = '#000'
  borderInner.LineStyle = Application.Enum.XlLineStyle.xlContinuous
  borderInner.Weight = Application.Enum.XlBorderWeight.xlThin
  // 设置区域外围边框
  const borderOut = await borders.Item(Application.Enum.XlBordersIndex.xlOutside)
  borderOut.Color = '#000'
  borderOut.LineStyle = Application.Enum.XlLineStyle.xlContinuous
  borderOut.Weight = Application.Enum.XlBorderWeight.xlThin
}

// 拆分工资表
async function splitData() {
  const headerInfo = await getHeaderInfo()
  console.log(headerInfo);
  const data = await getData()
  // 工作表对象
  const sheets = await Application.ActiveWorkbook.Sheets
  // 添加工作表
  await sheets.Add(null, null, 1, Application.Enum.XlSheetType.xlWorksheet, '个人工资条拆分表')

  let range = Application.Range(`A${headerRow + 1}`)
  for (let i = 0; i < data.length; i++) {
    const row = await range.Row
    await createHeader(row - headerRow, headerInfo)
    let cell = range
    for (let j = 0; j < totalLen; j++) {
      cell.Value = data[i][j]
      cell = await cell.Offset(0, 1)
    }
    await setBorder(row - headerRow)
    range = await range.Offset(headerRow + 2, 0)
  }
}

document.getElementById('start-split').addEventListener('click', splitData)
document.getElementById('template-create').addEventListener('click', createTemplate)
