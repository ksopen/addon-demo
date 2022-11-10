import { reactive } from 'vue'
// import './jssdk/index.js' jssdk需手动引入，jssdk下载地址参照readme
const jssdkApi = (window as any).jssdk.api
let app:any
let sheets:any
let luckySheet: any
const LotteryListName = '抽奖名单'
const prizeListName = '奖品池'
const luckyListName = '中奖名单'
const lotteryHeader = ['姓名', '工号']
const prizeHeader = ['奖品类型', '奖品名', '奖品数']
const LuckyHeader = ['奖品类型', '奖品名', '姓名', '工号']

interface IPerson {
  name: string,
  code: string,
}
interface IPrize {
  prizeType: string,
  prizeName: string,
  prizeNum: number
}

interface ILuckyPerson {
  name: string,
  code: string,
  prizeType: string,
  prizeName: string,
}

export const state: {
  personList: IPerson[],
  luckyList: ILuckyPerson[],
  unluckyList: IPerson[],
  prizeList: IPrize[],
  activePrizeIndex: number,
  activeLuckIndex: number,
  sheetsInfo: {
    name: string,
    sheet: any
  }[],
  activePersons: string[],
  activeTimes: number,
  status: string
} = reactive({
  personList: [],
  luckyList: [],
  unluckyList: [],
  prizeList: [],
  activePrizeIndex: 0,
  activeLuckIndex: 2,
  sheetsInfo: [],
  status: '',
  activePersons: [],
  activeTimes: -1,
})

export const numToAlphabet = (num: number) => {
  let str = ''
  while(num > 0) {
    let m = num % 26

    if (m === 0) {
      m = 26
    }

    str = String.fromCharCode(m + 64) + str
    num = (num - m ) / 26
  }

  return str
}

export const initApp = async () => {
  await jssdkApi.ready()
  app = jssdkApi.Application
  
  await getSheets()  
}

const getSheets = async () => {
  if (!app) {
    await initApp()
  }
  state.sheetsInfo = []
  sheets = await app.ActiveWorkbook.Sheets
  const sheetsCount = await sheets.Count
  for (let i =1; i <= sheetsCount; i++) {
    const sheet = await sheets.Item(i)
    const name = await sheet.Name
    state.sheetsInfo.push({
      sheet,
      name
    })
  }
}

const createSheetHeader = async (sheet: any, sheetHeader: string[]) => {
  sheetHeader.map(async (item, index) => {
    const range = await sheet.Range(`${numToAlphabet(index + 1)}1`)
    range.Value = item
  })
}

const initLotterySheet = async (sheetName: string, sheetHeader: string[]) => {
  if (!app) {
    await initApp()
  }

  let sheet
  state.sheetsInfo.map((item: any) => {
    if (item.name === sheetName) {
      sheet = item.sheet
    }
  })

  if (!sheet) {
    await sheets.Add({
      Before: 1,
      Name: sheetName
    })

    await getSheets()

    sheet = await sheets.Item(1)
    await createSheetHeader(sheet, sheetHeader)
  }

  return sheet
}

const getPersonList = async () => {
  const sheet = await initLotterySheet(LotteryListName, lotteryHeader)
  const list = []
  let isEnd = false
  let rowIndex = 2
  do {
    const name = await sheet.Range(`A${rowIndex}`).Text
    if (name.length === 0) { 
      isEnd = true; 
      break
    }
    const code = await sheet.Range(`B${rowIndex}`).Text
    list.push({ name, code })
    state.unluckyList.push({
      name,
      code
    })
    rowIndex++
  } while(!isEnd)
  state.personList = list
}

const getPrizeList = async () => {
  
  const sheet = await initLotterySheet(prizeListName, prizeHeader)
  const list = []
  let isEnd = false
  let rowIndex = 2
  do {
    const prizeType = await sheet.Range(`A${rowIndex}`).Text
    if (prizeType.length === 0) { 
      isEnd = true; 
      break
    }
    const prizeNumStr= await sheet.Range(`B${rowIndex}`).Text
    const prizeName = await sheet.Range(`C${rowIndex}`).Text
    const prizeNum = Number(prizeNumStr)
    list.push({ prizeType, prizeName, prizeNum })
    rowIndex++
  } while(!isEnd)
  state.prizeList = list
}

const getLuckyList = async () => {
  luckySheet = await initLotterySheet(luckyListName, LuckyHeader)
}

const activePersons = (activePersonNum: number, totalNum: number) => {
  const activePersonsIndex:number[] = []
  const activePersons = []
  while (activePersonsIndex.length < activePersonNum) {
    let matchIndex = Math.floor(Math.random() * totalNum)
    if (activePersonsIndex.indexOf(matchIndex) === -1) {
      activePersonsIndex.push(matchIndex)
      activePersons.push(state.unluckyList[matchIndex].name)
    }
  }
  state.activePersons = activePersons
  return activePersonsIndex
}

const startDraw = async () => {
  state.activeTimes++
  state.status = 'start'
  const activePrize = state.prizeList[state.activePrizeIndex]
  const activePersonNum = activePrize.prizeNum
  const totalNum = state.unluckyList.length
  let intervalCount = 0
  let activePersonsIndexs
  const interval = setInterval(() => {
    activePersonsIndexs = activePersons(activePersonNum, totalNum)
    intervalCount++
    if (intervalCount > 50) {
      clearInterval(interval)
      activePersonsIndexs.map((index: number) => {
        state.luckyList.push({
          ...state.unluckyList[index], 
          prizeType: activePrize.prizeType, 
          prizeName:activePrize.prizeName 
        })
      })
      state.status = 'end'
      state.activePrizeIndex++
      activePersonsIndexs.map((index: number) => {
        state.unluckyList.splice(index, 1)
      })
    }
  }, 100)
}

export const writeResult = async () => {
  state.luckyList.map(async (item, index) => {
    const prizeType = await luckySheet.Range(`A${index+2}`)
    prizeType.Value = item.prizeType
    const prizeName = await prizeType.Offset(0, 1)
    prizeName.Value = item.prizeName
    const name = await prizeName.Offset(0, 1)
    name.Value = item.name
    const code = await name.Offset(0, 1)
    code.Value = item.code
  })
}

export const onDrawLottery = async () => {
  if (state.personList.length === 0) {
    await getLuckyList()
    await getPrizeList()
    await getPersonList() 
  }
  startDraw()
}
