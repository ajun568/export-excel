![banner.png](https://p3-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/1dba78d5da6f49b4a6d4ff80f6c3c764~tplv-k3u1fbpfcp-watermark.image)

> Hello, 各位勇敢的小伙伴, 大家好, 我是你们的嘴强王者小五, 身体健康, 脑子没病.
>
> 本人有丰富的脱发技巧, 能让你一跃成为资深大咖.
>
> 一看就会一写就废是本人的主旨, 菜到抠脚是本人的特点, 卑微中透着一丝丝刚强, 傻人有傻福是对我最大的安慰.
>
> 欢迎来到`小五`的`随笔系列`之`前端导出Excel在线指北`.

## 启动项目

* yarn install

* yarn start

## 写在前面

**双手奉上代码链接** [传送门 - ajun568](https://github.com/ajun568/export-excel)

**双脚奉上最终效果图**

![excel1.png](https://p1-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/89c83c07ce6043c38c0de11cb283a0a6~tplv-k3u1fbpfcp-watermark.image)

**观前提醒**

👺 本文最终实现效果如上图, 具体功能为: `导出Excel + 多个Sheet + 可合并的多行表头`. 代码部分采用 `React+TS` 作为工具进行编写.

## 准备工作

👺 **安装 xlsx.js** `npm install xlsx`

👺 **写入Excel文件:** `XLSX.write(workbook, writeOpts)`

**workbook 👇**

* SheetNames `@types string[]`: 当前 Sheet 的名称

* Sheets: 当前sheet的对象, 格式如下

```js
[SheetNames]: {
  "!refs": "A1:G7", // 表示从 第1行第A列 到 第7行第G列
  "!cols": [{wpx: 80} ... ], // 表示 列宽 80px
  "!rows": [{hpx: 20} ... ], // 表示 行高 20px
  "!merges": [{s: {r: 0, c: 2}, e: {r: 0, c: 3}} ... ], // 表示 将 第0行第2列 和 第0行第3列 进行合并 (s: start, e: end, c: column, r: row)
  "A1": {v: "姓名"}, // 表示第1行第A列 显示数据为 "姓名", 以此类推 ...
  ...
}
```

**writeOpts 👇**

```js
{
  type, // 数据编码, 本文采用 binary 二进制格式
  bookType, // 导出类型, 本文采用 xlsx 类型
  compression, // 是否使用 Gzip 压缩
}
```

## 下载文件

想要下载文件, 我小A第一个表示不服, 申请出战 <[a 标签的 download 属性](https://www.w3school.com.cn/tags/att_a_download.asp)>


![excel5.jpg](https://p9-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/cbe7f302196f41bf814961b23dee182d~tplv-k3u1fbpfcp-watermark.image)

通过 [URL.createObjectURL(Object)](https://developer.mozilla.org/zh-CN/docs/Web/API/URL/createObjectURL) 来创建下载所需的 URL. 由于每次调用都会产生新的 URL 对象, 故使用后记得释放, 释放方法 [URL.revokeObjectURL(FileUrl)](https://developer.mozilla.org/zh-CN/docs/Web/API/URL/revokeObjectURL)

通过模拟 click 事件触发 a 标签, 以实现下载

```js
const saveAs = (obj: Blob, fileName?: string): void => {
  const temp = document.createElement('a')
  temp.download = fileName || 'download'
  temp.href = URL.createObjectURL(obj)
  temp.click()
  setTimeout(() =>  { URL.revokeObjectURL(temp.href) }, 100)
}
```

## 头部处理

**Mock数据**: 详细数据请跳转 [Github](https://github.com/ajun568/export-excel), 在 `mock.ts` 中查看

Header 部分数据格式

```js
[
  ...
  {
    key: 'animal',
    value: '动物',
    child: [
      {
        key: 'dog',
        value: '狗',
        child: [
          {
            key: 'corgi',
            value: '柯基',
          },
          {
            key: 'husky',
            value: '哈士奇',
          },
        ],
      },
      {
        key: 'tiger',
        value: '老虎',
      },
    ],
  },
  ...
]
```

Data 部分数据格式

```js
[
  {
    name: '黄刀小五',
    desc: '基于搜索引擎的复制粘贴攻城狮',
    watermelon: '喜欢',
    banana: '不喜欢',
    corgi: '喜欢',
    husky: '喜欢',
    tiger: '不喜欢',
  },
  ...
]
```

### 头部数据处理

👺 **分析**

* **Header** 数据为树形结构, 其深度为头部所占行数

* **Header** 数据要转换成 **Data** 数据的格式, 并与 **Data** 数组合并, 共同处理成导出所需格式

* 转换对象的 **key** 应为最小叶子结点的 **key**

* 转换对象的 **value** 应为当前层级的 **value** ( 即导出后当前行所显示的 **value** )

* 既然是树, 果断递归, 准没错

**🧟‍♂️ Code**

![excel2.png](https://p1-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/0ad332b6ae7c43c79b6c0957fd3e3f39~tplv-k3u1fbpfcp-watermark.image)

**🧟‍♂️ Image**

![excel3.png](https://p6-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/16d81953128b4c8c895bdd227134d091~tplv-k3u1fbpfcp-watermark.image)

### Merged 数据

```js
{
  s: { // start
    r: x, // row
    c: y, // column
  },
  e: { ... } // end
}
```

👺 **分析**

* 将处理后的头部数据看成一个矩阵

* 行或列中, 相邻元素若相同, 则进行合并

**tips:** 本文采用的是判断相邻 value 值是否相等进行合并, 若有需求, 建议改写为对象形式加以完善.

**🧟‍♂️ Code**

![excel4.png](https://p1-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/f70c80a4ace94a249c7762dfbcb7a6f8~tplv-k3u1fbpfcp-watermark.image)

**🧟‍♂️ Image**

![excel6.png](https://p1-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/c54922a82130485ba41d190383d1b80c~tplv-k3u1fbpfcp-watermark.image)

## 生成sheet数据

* 利用`Object.assign`进行对象合并

* 利用`String.fromCharCode(65 + i)`对列进行大写字母的转换

**🧟‍♂️ Code**


![excel13.png](https://p6-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/3d78ca2343b141859f50c42cff83e769~tplv-k3u1fbpfcp-watermark.image)

**🧟‍♂️ Image**

![excel12.png](https://p6-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/615fcc9ed84a4aebb1428c768efd44d0~tplv-k3u1fbpfcp-watermark.image)

## 转换字节流

利用 [new ArrayBuffer(str)](https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects/ArrayBuffer) 创建一个缓冲区, 使用 `new Uint8Array(buf)` 引用


![excel8.gif](https://p6-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/21d0f95481a5450e832c9d7ad64d464d~tplv-k3u1fbpfcp-watermark.image)

因为 **unicode** 编码是 **0~65535**, 而 **Uint8Array** 范围为 **0~255**, 故需要按位与 **0xFF**, 以保持位数一致

```js
const s2ab = (str: string): ArrayBuffer => {
  let buf = new ArrayBuffer(str.length)
  let view = new Uint8Array(buf)

  for (let i = 0; i !== str.length; ++i) {
    view[i] = str.charCodeAt(i) & 0xFF
  }

  return buf
}
```

## 导出文件

结合前文 **准备工作** 部分所讲, 导出的代码逻辑就出来了, 直接上代码

![excel14.png](https://p3-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/3230a2cbcdf342d18a927dfbc0f41880~tplv-k3u1fbpfcp-watermark.image)

## 结束语

开源版本不支持设置样式, 若有需求, 可采用 **付费版本** 或使用 `xlsx-style`, 使用方法与本文一致. 大家可参照[文档](https://github.com/protobi/js-xlsx)自行添加样式部分.


![excel7.gif](https://p9-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/2e7074cb4fd04354a389f20470dd2c15~tplv-k3u1fbpfcp-watermark.image)


## 参考🔗链接

[[Github] SheetJS ~ js-xlsx](https://github.com/SheetJS/sheetjs#writing-options)

[[mySoul] 优雅 | 前后端优雅的导入导出Excel](https://juejin.cn/post/6872375842358919175)

[[Seefly] 前端使用xlsx.js导出有复杂表头的excel](http://www.seefly.top/archives/%E5%89%8D%E7%AB%AF%E4%BD%BF%E7%94%A8xlsxjs%E5%AF%BC%E5%87%BA%E6%9C%89%E5%A4%8D%E6%9D%82%E8%A1%A8%E5%A4%B4%E7%9A%84excelmd#6%E5%8D%95%E5%85%83%E6%A0%BC%E5%88%97%E5%AE%BD)
