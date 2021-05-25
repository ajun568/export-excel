import * as XLSX from 'xlsx'

export interface SheetProps {
  ento: never;
  data: never[];
  sheetName: string;
}

interface TreeProps {
  child?: TreeProps[],
  [key: string]: any,
}

interface ObjectProps {
  [key: string]: string,
}

interface MergedItemProps {
  r: number,
  c: number,
}

interface MergedProps {
  e: MergedItemProps,
  s: MergedItemProps,
}

interface MergedRowProps {
  [key: string]: MergedProps,
}

interface TempProps {
  i?: number,
  value?: string,
}

interface EntoMapProps {
  key: string,
  value: string,
  child?: EntoMapProps[],
}

// 字符串转字节流
const s2ab = (str: string): ArrayBuffer => {
  let buf = new ArrayBuffer(str.length)
  let view = new Uint8Array(buf)

  for (let i = 0; i !== str.length; ++i) {
    view[i] = str.charCodeAt(i) & 0xFF
  }

  return buf
}

// 查找树的所有叶子节点
const getAllLeaf = (tree: TreeProps[]): string[] => {
  let result: string[] = []

  const getLeaf = (tree: TreeProps[]): void => {
    tree.forEach(item => {
      if (!item.child) {
        if (item.key) result.push(item.key)
      } else {
        getLeaf(item.child)
      }
    })
  }
  
  getLeaf(tree)
  return result
}

// 生成头部 Merged 数组
const createMerged = (arr: ObjectProps[]): MergedProps[] => {
  let rowArr: MergedRowProps[] = []
  let columnArr: string[][] = []
  let mergeArr: MergedProps[] = []

  arr.forEach((item, i) => {
    let temp: TempProps = {}
    let index = 0
    rowArr[i] = {}

    for (let key in item) {
      if (temp.value === item[key]) {
        if (rowArr[i][item[key]]) {
          rowArr[i][item[key]].e = { r: i, c: index }
        } else {
          rowArr[i][item[key]] = { s: { r: i, c: index - 1 }, e: { r: i, c: index }}
        }
      }

      if (i === 0) columnArr[index] = []
      columnArr[index][i] = item[key]

      index += 1
      temp['i'] = index
      temp['value'] = item[key]
    }
  })

  rowArr.forEach(item => {
    if (JSON.stringify(item) === '{}') return
    mergeArr.push(...Object.values(item))
  })

  columnArr.forEach((child, childIndex) => {
    let temp = {
      value: '',
      time: 1,
      start: 0,
      end: 0,
    }

    child.forEach((item, i) => {
      if (i === 0) {
        temp.value = item
        return
      }

      if (item === temp.value) {
        temp.time += 1
        temp.end = i

        if (i === child.length - 1 && temp.time > 1) {
          mergeArr.push({ s: { r: temp.start, c: childIndex }, e: { r: temp.end, c: childIndex }})
        }
      } else {
        if (temp.time > 1) {
          mergeArr.push({ s: { r: temp.start, c: childIndex }, e: { r: temp.end, c: childIndex }})
        }

        temp.value = item
        temp.time = 1
        temp.start = i
        temp.end = i
      }
    })
  })

  return mergeArr
}

// 处理头部数据信息
const dealHeaderData = (arr: EntoMapProps[]) => {
  const recursiveTree = (
    arr: EntoMapProps[], parentData: TreeProps[] = [], res: ObjectProps[] = [], level: number = 0
  ) => {
    arr.forEach(item => {
      if (!res[level]) res[level] = {}

      if (item.child) {
        if (!parentData[level]) parentData[level] = []
        parentData[level].push({...item, leaf: getAllLeaf(item.child)})
        res = recursiveTree(item.child, parentData, res, level + 1)
      } else {
        if (item.key) res[level][item.key] = item.value
      }

      for (let i = 0; i <= level; i++) {
        if (parentData[i]) {
          parentData[i].forEach((child: TreeProps) => {
            if (!res[i][item.key] && child.leaf.includes(item.key)) {
              res[i][item.key] = child.value
            }
          })
        }
      }
    })
    return res
  }

  const dealTree = recursiveTree(arr)
  let temp = {}
  const headerArr = dealTree.map((item, index) => {
    let itemTemp = index ? { ...temp, ...item } : item
    temp = itemTemp
    return itemTemp
  })
  const mergeArr = createMerged(headerArr)

  return { headerArr, mergeArr }
}

// 处理sheet内容
const dealSheet = (
  data: never[], ento: never[], cols?: number[]
) => {
  const { headerArr, mergeArr } = dealHeaderData(ento)
  if (!headerArr.length) {
    console.warn('处理sheet的参数存在错误, 请检查!')
    return
  }

  const entoMapObj = headerArr[headerArr.length - 1]
  const json = [...headerArr, ...data]
  let jsonData: any = {}

  const dealJson = json.map((v, i) => 
    Object.keys(entoMapObj).map((k, j) => 
      Object.assign({}, {
        v: v[k],
        position: String.fromCharCode(65 + j) + (i + 1),
      })
    )
  )

  dealJson.reduce((prev, next) => 
    prev.concat(next)).forEach(v => 
    jsonData[v.position] = { v: v.v })

  const outputPos = Object.keys(jsonData)
  const colsArr = cols?.map(item => {
    return { wpx: item }
  })

  return Object.assign(
    {},
    jsonData,
    { '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] },
    { '!rows': dealJson.map(() => { return { hpx: 20 } })},
    cols ? { '!cols': colsArr } : {},
    { '!merges': mergeArr },
  )
}

// 导出
const saveAs = (obj: Blob, fileName?: string): void => {
  const temp = document.createElement('a')
  temp.download = fileName || 'download'
  temp.href = URL.createObjectURL(obj)
  temp.click()
  setTimeout(() =>  { URL.revokeObjectURL(temp.href) }, 100)
}

// 下载
const exportExcel = (
  data: SheetProps[], fileName: string, opts?: XLSX.WritingOptions & { cols?: number[] } | undefined
): void => {
  let sheetNames: string[] = []
  let sheets: XLSX.WorkSheet = {}

  data.forEach((item, i) => {
    let sheetName = item?.sheetName || `sheet${i ? i + 1 : ''}`
    sheetNames.push(sheetName)
    sheets[sheetName] = dealSheet(item.data, item.ento, opts?.cols)
  })

  const writeOpts = {
    type: 'binary' as const,
    bookSST: false,
    bookType: opts?.bookType || 'xlsx',
    sheet: opts?.sheet || '',
    compression: opts?.compression || false,
  }

  const workbook = {
    SheetNames: sheetNames,
    Sheets: sheets,
  }

  const saveData = new Blob(
    [ s2ab(XLSX.write(workbook, writeOpts)) ],
    { type: '' },
  )

  saveAs(saveData, fileName)
}

export default exportExcel
