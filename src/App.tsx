import React from 'react'
import { ento1, data1, ento2, data2 } from './mock'
import exportExcel, { SheetProps } from './export'

const App: React.FC = () => {
  const toExport = () => {
    console.log('这里是【 调查问卷 】头部信息: ', ento1)
    console.log('这里是【 调查问卷 】数据信息: ', data1)
    console.log('这里是【 个人信息 】头部信息: ', ento2)
    console.log('这里是【 个人信息 】数据信息: ', data2)

    const sheet1 = { ento: ento1, data: data1, sheetName: '调查问卷' }
    const sheet2 = { ento: ento2, data: data2, sheetName: '个人信息' }
    
    exportExcel(
      [sheet1, sheet2] as SheetProps[],
      '黄刀小五.xlsx',
      { cols: [80, 200, 80, 80, 80, 80, 80] }
    )
  }

  return (
    <main style={{ padding: 50 }}>
      <button onClick={toExport}>点击这个Button会触发导出效果</button>
    </main>
  )
}

export default App
