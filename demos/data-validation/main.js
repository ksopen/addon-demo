import './app-jssdk/index.js'

const jssdkApi = window.jssdk.api
const id2province = {
  '11': '北京市',
  '12': '天津市',
  '13': '河北省',
  '14': '山西省',
  '15': '内蒙古自治区',
  '21': '辽宁省',
  '22': '吉林省',
  '23': '黑龙江省',
  '31': '上海市',
  '32': '江苏省',
  '33': '浙江省',
  '34': '安徽省',
  '35': '福建省',
  '36': '江西省',
  '37': '山东省',
  '41': '河南省',
  '42': '湖北省',
  '43': '湖南省',
  '44': '广东省',
  '45': '广西壮族自治区',
  '46': '海南省',
  '50': '重庆市',
  '51': '四川省',
  '52': '贵州省',
  '53': '云南省',
  '54': '西藏自治区',
  '61': '陕西省',
  '62': '甘肃省',
  '63': '青海省',
  '64': '宁夏回族自治区',
  '65': '新疆维吾尔自治区',
}

// 只对18位身份证进行检验
function parseIdentify(id) {
  id = id.toLowerCase()
  if (id.length != 18) {
    return false
  }
  let result = {}
  let province = id2province[id.substr(0, 2)]
  if (province === undefined) {
    return false
  }
  result.province = province

  if (parseInt(id[16]) % 2 == 0) {
    result.sex = '女'
  } else {
    result.sex = '男'
  }

  result.birthday = id.substr(6, 4) + '-' + id.substr(10, 2) + '-' + id.substr(12, 2)


  const factors = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
  const checksums = ['1', '0', 'x', '9', '8', '7', '6', '5', '4', '3', '2']
  let sum = 0
  for (let i = 0; i < 17; i++) {
    sum += parseInt(id[i]) * factors[i]
  }
  let checksum = checksums[sum % 11]
  if (checksum !== id[17]) {
    return false
  }

  return result
}

async function verify() {
  const instance = await jssdkApi.ready()
  const Application = jssdkApi.Application
  if (!instance) {
    return
  }
  if (await Application.Selection.Columns.Count > 1) {
    alert('请选择一列')
    return
  }
  let selectionAddr = await Application.Selection.Address()

  //在身份证这一列后面, 插入3列: '省份', '性别', '生日'
  let newColumn = await Application.Selection.EntireColumn
  const titles = ['省份', '性别', '生日']
  let newColumns = []
  for (let i = 0; i < titles.length; i++) {
    newColumn = await newColumn.Offset(0, 1)
    await newColumn.Insert()
    const item = await newColumn.Item(1)
    item.Value = titles[i]
    newColumns.push(newColumn)
  }

  //逐行校验身份证的合法性, 提出身份信息, 并写入到新的列中
  let selection = Application.Range(selectionAddr)
  const count = await selection.Count
  for (let i = 1; i <= count; i++) {
    let item = await selection.Item(i)
    let id = await item.Text
    if (id === '') {
      break
    }

    let result = parseIdentify(id)
    if (result === false) {
      //身份证校验不通过, 标记为黄色
      console.info(await item.Address(), 'invalid', id)
      const interior = await item.Interior
      interior.Color = '#ffff00'
    } else {
      console.info(await item.Address(), 'valid', id, result)
      //省份
      const province = await newColumns[0].Item(i + 1)
      province.Value = result.province
      //性别
      const sex = await newColumns[1].Item(i + 1)
      sex.Value = result.sex
      //生日
      const birthday = await newColumns[2].Item(i + 1)
      birthday.Value = result.birthday
    }
  }
  await Application.WhenStacksEmpty()
}

document.getElementById('verify-button').addEventListener('click', verify)