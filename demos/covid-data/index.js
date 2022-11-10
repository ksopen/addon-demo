import './jssdk/dist/index.js'
import './style.css'

const jssdkApi = window.jssdk.api
let Application
(async () => {
  await jssdkApi.ready()
  Application = jssdkApi.Application
})()
const infoTitle = [
  ['现有确诊', '无症状', '现有疑似', '现有重症'],
  ['累计确诊', '境外输入', '累计治愈', '累计死亡']
]
const caseListTitle = ['疫情地区', '新增', '现有', '累计', '治愈', '死亡']
const topAddTitle = ['省市', '本土输入', '境外输入']

// 初始化疫情数据表格
async function initTable() {
  setStartTips(0)
  // 工作表对象
  const sheets = await Application.ActiveWorkbook.Sheets
  // 添加工作表
  await sheets.Add(null, null, 1, Application.Enum.XlSheetType.xlWorksheet, '全国疫情数据')

  // 全局设置样式
  // 活动工作簿中的活动工作表
  const activeSheet = await Application.ActiveWorkbook.ActiveSheet

  // 工作表上的所有行
  const rows = await activeSheet.Rows
  // 字体对象
  const allFont = await rows.Font
  // 设置字体大小
  allFont.Size = 12
  // 设置文本水平居中
  rows.HorizontalAlignment = await Application.Enum.XlVAlign.xlVAlignCenter
  // 设置单元格宽度
  rows.ColumnWidth = 120


  // 设置全国数据展示区域的样式
  const topArea = Application.Range('A2:G8')
  const topInterior = await topArea.Interior
  topInterior.Color = "#f5f6f7"
  let rowCell = Application.Range('A2')
  for (let i = 0; i < infoTitle.length; i++) {
    let columnCell = rowCell
    for (let j = 0; j < infoTitle[i].length; j++) {
      columnCell.Value = infoTitle[i][j]
      // 操作区域向下移动一个单位
      columnCell = await columnCell.Offset(1, 0)
      columnCell.Select()
      let font = await columnCell.Font
      font.Color = '#00bec9'
      font.Bold = true

      columnCell = await columnCell.Offset(1, 0)
      let fontnext = await columnCell.Font
      fontnext.Color = '#eb3941'

      // 移动到下一列的初始位置
      columnCell = await columnCell.Offset(-2, 2)
    }
    // 移动到下一个样式格式化的区域
    rowCell = await rowCell.Offset(4, 0)
  }

  // 对全国各省数据展示区域进行表格样式设置
  const nextAreaTitle = Application.Range('A10:F10')
  const font = await nextAreaTitle.Font
  font.Bold = true
  const count = await nextAreaTitle.Count
  // 表头
  for (let i = 0; i < count; i++) {
    const item = await nextAreaTitle.Item(1, i + 1)
    item.Value = caseListTitle[i]
  }
  // 内容区域
  for (let i = 0; i < 34; i++) {
    let contentCell = Application.Range(`A${11 + i}`)
    const interior = await contentCell.Interior
    const cellFont = await contentCell.Font
    cellFont.Color = '#ffffff'
    interior.Color = '#00bec9'
    let contentRow = Application.Range(`B${11 + i}:F${11 + i}`)
    const rowInterior = await contentRow.Interior
    rowInterior.Color = '#f5f6f7'
  }

  const addTopArea = Application.Range('A47:C47')
  await addTopArea.Merge()
  addTopArea.Value = '新增确诊分布Top'
  const titleFont = await addTopArea.Font
  titleFont.Bold = true
  let topAreaTitle = Application.Range('A48')
  for (let i = 0; i < topAddTitle.length; i++) {
    topAreaTitle.Value = topAddTitle[i]
    topAreaTitle = await topAreaTitle.Offset(0, 1)
  }

  await writeData()
  setEndTips(0)
}
const DATA_TYPE = ['curConfirm', 'asymptomatic', 'unconfirmed', 'icu', 'confirmed', 'overseasInput', 'cured', 'died']

// 更新数据展示
async function writeData() {
  const data = await getData()
  const timeArea = Application.Range('A1:F1')
  await timeArea.Merge()
  timeArea.Value = `更新时间：${data.mapLastUpdatedTime}`

  const summaryDataIn = data.summaryDataIn
  let rowCell = Application.Range('A3')
  for (let i = 0; i < infoTitle.length; i++) {
    let columnCell = rowCell
    for (let j = 0; j < infoTitle[i].length; j++) {
      const flag = DATA_TYPE[i * 4 + j]
      columnCell.Value = summaryDataIn[flag]
      columnCell = await columnCell.Offset(1, 0)
      const numstr = summaryDataIn[`${flag}Relative`] > 0 ? `+${summaryDataIn[`${flag}Relative`]}` : `${summaryDataIn[`${flag}Relative`]}`
      columnCell.Value = `昨日 ${numstr}`
      columnCell = await columnCell.Offset(-1, 2)
    }
    rowCell = await rowCell.Offset(4, 0)
  }

  const caseList = data.caseList
  for (let i = 0; i < caseList.length; i++) {
    let caseCell = Application.Range(`A${11 + i}`)
    caseCell.Value = caseList[i].area

    caseCell = await caseCell.Offset(0, 1)
    caseCell.Value = caseList[i].confirmedRelative

    caseCell = await caseCell.Offset(0, 1)
    caseCell.Value = caseList[i].curConfirm

    caseCell = await caseCell.Offset(0, 1)
    caseCell.Value = caseList[i].confirmed

    caseCell = await caseCell.Offset(0, 1)
    caseCell.Value = caseList[i].crued

    caseCell = await caseCell.Offset(0, 1)
    caseCell.Value = caseList[i].died
  }
  const newAddTopProvince = data.newAddTopProvince
  for (let i = 0; i < newAddTopProvince.length; i++) {
    let topCell = Application.Range(`A${49 + i}`)
    topCell.Value = newAddTopProvince[i].name + ''

    topCell = await topCell.Offset(0, 1)
    topCell.Value = newAddTopProvince[i].local + ''

    topCell = await topCell.Offset(0, 1)
    topCell.Value = newAddTopProvince[i].overseasInput + ''
  }
}

// 防疫数据获取
async function getData() {
  const covidRes = await fetch('https://team-app-cdn.wpscdn.cn/plugin-game-demo/covid19.json')
  const covid = await covidRes.json()
  console.log('covid',covid)
  return {
    caseList: covid.caseList,
    mapLastUpdatedTime: covid.mapLastUpdatedTime,
    summaryDataIn: covid.summaryDataIn,
    newAddTopProvince: covid.newAddTopProvince
  }
}

function setStartTips(index) {
  const el = document.getElementsByClassName('tips')[index]
  el.style.color = 'red'
  el.innerText = '正在加载生成中，请等待'
}

function setEndTips(index) {
  const el = document.getElementsByClassName('tips')[index]
  el.style.color = 'green'
  el.innerText = '已完成操作'
}

async function createChart() {
  setStartTips(1)
  // 活动工作簿中的活动工作表
  var activeSheet = Application.ActiveWorkbook.ActiveSheet
  // 当前工作表上的所有 Shape 对象的集合
  var shapes = activeSheet.Shapes
  // 创建图表
  const shape = shapes.AddChart2(340, Application.Enum.XlChartType.xlColumnClustered, 700, 300, 600, 400)
  const chart = shape.Chart
  // 选择数据区域
  const source = activeSheet.Range('A48:C58')
  
  await source.Select()
  setEndTips(1)
  await chart.SetSourceData(source,2)
  chart.HasTitle = true
  chart.ChartTitle.Text = '疫情新增top城市'
  
}

document.getElementById('init').addEventListener('click', initTable)
document.getElementById('chart').addEventListener('click', createChart)
