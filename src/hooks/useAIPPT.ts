import { ref } from 'vue'
import { nanoid } from 'nanoid'
import type { ImageClipDataRange, PPTElement, PPTImageElement, PPTShapeElement, PPTTableElement, PPTTextElement, Slide, TableCell, TableCellStyle, TextAlign, TextType } from '@/types/slides'
import type { AIPPTSlide } from '@/types/AIPPT'
import { useSlidesStore } from '@/store'
import useAddSlidesOrElements from './useAddSlidesOrElements'
import useSlideHandler from './useSlideHandler'

interface ImgPoolItem {
  id: string
  src: string
  width: number
  height: number
}

export default () => {
  const slidesStore = useSlidesStore()
  const { addSlidesFromData } = useAddSlidesOrElements()
  const { isEmptySlide } = useSlideHandler()

  const imgPool = ref<ImgPoolItem[]>([])
  const transitionIndex = ref(0)
  const transitionTemplate = ref<Slide | null>(null)

  const checkTextType = (el: PPTElement, type: TextType) => {
    return (el.type === 'text' && el.textType === type) || (el.type === 'shape' && el.text && el.text.type === type)
  }
  
  const getUseableTemplates = (templates: Slide[], n: number, type: TextType) => {
    if (n === 1) {
      const list = templates.filter(slide => {
        const items = slide.elements.filter(el => checkTextType(el, type))
        const titles = slide.elements.filter(el => checkTextType(el, 'title'))
        const texts = slide.elements.filter(el => checkTextType(el, 'content'))
  
        return !items.length && titles.length === 1 && texts.length === 1
      })
  
      if (list.length) return list
    }
  
    let target: Slide | null = null
  
    const list = templates.filter(slide => {
      const len = slide.elements.filter(el => checkTextType(el, type)).length
      return len >= n
    })
    if (list.length === 0) {
      const sorted = templates.sort((a, b) => {
        const aLen = a.elements.filter(el => checkTextType(el, type)).length
        const bLen = b.elements.filter(el => checkTextType(el, type)).length
        return aLen - bLen
      })
      target = sorted[sorted.length - 1]
    }
    else {
      target = list.reduce((closest, current) => {
        const currentLen = current.elements.filter(el => checkTextType(el, type)).length
        const closestLen = closest.elements.filter(el => checkTextType(el, type)).length
        return (currentLen - n) <= (closestLen - n) ? current : closest
      })
    }
  
    return templates.filter(slide => {
      const len = slide.elements.filter(el => checkTextType(el, type)).length
      const targetLen = target!.elements.filter(el => checkTextType(el, type)).length
      return len === targetLen
    })
  }
  
  const getAdaptedFontsize = ({
    text,
    fontSize,
    fontFamily,
    width,
    maxLine,
  }: {
    text: string
    fontSize: number
    fontFamily: string
    width: number
    maxLine: number
  }) => {
    const canvas = document.createElement('canvas')
    const context = canvas.getContext('2d')!
  
    let newFontSize = fontSize
    const minFontSize = 10
  
    while (newFontSize >= minFontSize) {
      context.font = `${newFontSize}px ${fontFamily}`
      const textWidth = context.measureText(text).width
      const line = Math.ceil(textWidth / width)
  
      if (line <= maxLine) return newFontSize
  
      const step = newFontSize <= 22 ? 1 : 2
      newFontSize = newFontSize - step
    }
  
    return minFontSize
  }
  
  const getFontInfo = (htmlString: string) => {
    const fontSizeRegex = /font-size:\s*(\d+(?:\.\d+)?)\s*px/i
    const fontFamilyRegex = /font-family:\s*['"]?([^'";]+)['"]?\s*(?=;|>|$)/i
  
    const defaultInfo = {
      fontSize: 16,
      fontFamily: 'Microsoft Yahei',
    }
  
    const fontSizeMatch = htmlString.match(fontSizeRegex)
    const fontFamilyMatch = htmlString.match(fontFamilyRegex)
  
    return {
      fontSize: fontSizeMatch ? (+fontSizeMatch[1].trim()) : defaultInfo.fontSize,
      fontFamily: fontFamilyMatch ? fontFamilyMatch[1].trim() : defaultInfo.fontFamily,
    }
  }
  
  const getNewTextElement = ({
    el,
    text,
    maxLine,
    longestText,
    digitPadding,
  }: {
    el: PPTTextElement | PPTShapeElement
    text: string
    maxLine: number
    longestText?: string
    digitPadding?: boolean
  }): PPTTextElement | PPTShapeElement => {
    const padding = 10
    const width = el.width - padding * 2 - 2
  
    let content = el.type === 'text' ? el.content : el.text!.content
  
    const fontInfo = getFontInfo(content)
    const size = getAdaptedFontsize({
      text: longestText || text,
      fontSize: fontInfo.fontSize,
      fontFamily: fontInfo.fontFamily,
      width,
      maxLine,
    })
  
    const parser = new DOMParser()
    const doc = parser.parseFromString(content, 'text/html')
  
    const treeWalker = document.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT)
  
    const firstTextNode = treeWalker.nextNode()
    if (firstTextNode) {
      if (digitPadding && firstTextNode.textContent && firstTextNode.textContent.length === 2 && text.length === 1) {
        firstTextNode.textContent = '0' + text
      }
      else firstTextNode.textContent = text
    }
  
    if (doc.body.innerHTML.indexOf('font-size') === -1) {
      const p = doc.querySelector('p')
      if (p) p.style.fontSize = '16px'
    }
  
    content = doc.body.innerHTML.replace(/font-size:(.+?)px/g, `font-size: ${size}px`)
  
    return el.type === 'text' ? { ...el, content, lineHeight: size < 15 ? 1.2 : el.lineHeight } : { ...el, text: { ...el.text!, content } }
  }

  const getUseableImage = (el: PPTImageElement): ImgPoolItem | null => {
    let img: ImgPoolItem | null = null
  
    let imgs = []
  
    if (el.width === el.height) imgs = imgPool.value.filter(img => img.width === img.height)
    else if (el.width > el.height) imgs = imgPool.value.filter(img => img.width > img.height)
    else imgs = imgPool.value.filter(img => img.width <= img.height)
    if (!imgs.length) imgs = imgPool.value
  
    img = imgs[Math.floor(Math.random() * imgs.length)]
    imgPool.value = imgPool.value.filter(item => item.id !== img!.id)
  
    return img
  }
  
  const getNewImgElement = (el: PPTImageElement): PPTImageElement => {
    const img = getUseableImage(el)
    if (!img) return el
  
    let scale = 1
    let w = el.width
    let h = el.height
    let range: ImageClipDataRange = [[0, 0], [0, 0]]
    const radio = el.width / el.height
    if (img.width / img.height >= radio) {
      scale = img.height / el.height
      w = img.width / scale
      const diff = (w - el.width) / 2 / w * 100
      range = [[diff, 0], [100 - diff, 100]]
    }
    else {
      scale = img.width / el.width
      h = img.height / scale
      const diff = (h - el.height) / 2 / h * 100
      range = [[0, diff], [100, 100 - diff]]
    }
    const clipShape = (el.clip && el.clip.shape) ? el.clip.shape : 'rect'
    const clip = { range, shape: clipShape }
    const src = img.src
  
    return { ...el, src, clip }
  }
  
  const getMdContent = (content: string) => {
    const regex = /```markdown([^```]*)```/
    const match = content.match(regex)
    if (match) return match[1].trim()
    return content.replace('```markdown', '').replace('```', '')
  }
  
  const getJSONContent = (content: string) => {
    const regex = /```json([^```]*)```/
    const match = content.match(regex)
    if (match) return match[1].trim()
    return content.replace('```json', '').replace('```', '')
  }

  const presetImgPool = (imgs: ImgPoolItem[]) => {
    imgPool.value = imgs
  }
  const parseHTMLTable = (htmlString: string): PPTTableElement | null => {
    try {
      const parser = new DOMParser()
      const doc = parser.parseFromString(htmlString, 'text/html')
      const table = doc.querySelector('table')
      
      if (!table) return null

      const rows = Array.from(table.querySelectorAll('tr'))
      if (rows.length === 0) return null

      // 用于跟踪每个位置的占用情况
      const occupiedMatrix: Array<Array<TableCell | null>> = []
      
      let maxCols = 0
      const defaultRowHeight = 36

      // 颜色值清理函数
      const cleanColorValue = (colorStr: string): string => {
        return colorStr.replace(/;;+/g, ';').replace(/;\$/, '').trim()
      }

      // 第一遍：解析所有单元格并填充占用矩阵
      rows.forEach((row, rowIndex) => {
        const cells = Array.from(row.querySelectorAll('td, th'))
        let colIndex = 0
        
        // 确保占用矩阵有足够的行
        while (occupiedMatrix.length <= rowIndex) {
          occupiedMatrix.push([])
        }

        cells.forEach(cell => {
          // 找到下一个未被占用的列位置
          while (occupiedMatrix[rowIndex][colIndex] !== undefined) {
            colIndex++
          }

          const colspan = parseInt(cell.getAttribute('colspan') || '1')
          const rowspan = parseInt(cell.getAttribute('rowspan') || '1')
          
          // 获取单元格文本内容
          const dataValue = cell.getAttribute('data-value')
          let cellText = dataValue || cell.textContent?.trim() || ''
          
          // 从嵌套的div中提取文本（如果存在）
          if (!cellText) {
            const cellTextDiv = cell.querySelector('.cell-text')
            if (cellTextDiv) {
              cellText = cellTextDiv.textContent?.trim() || ''
            }
          }
          
          // 解析样式
          const inlineStyle = cell.getAttribute('style') || ''
          const cellStyle: TableCellStyle = {}
          
          // 解析各种样式属性
          const backgroundColorMatch = inlineStyle.match(/background-color:\s*([^;]+)/i)
          if (backgroundColorMatch) {
            const bgColor = cleanColorValue(backgroundColorMatch[1])
            if (bgColor && bgColor !== 'transparent' && bgColor !== 'rgba(0, 0, 0, 0)') {
              cellStyle.backcolor = bgColor
            }
          }
          
          const colorMatch = inlineStyle.match(/(?:^|;)\s*color:\s*([^;]+)/i)
          if (colorMatch) {
            const textColor = cleanColorValue(colorMatch[1])
            if (textColor && textColor !== 'transparent') {
              cellStyle.color = textColor
            }
          }
          
          const fontSizeMatch = inlineStyle.match(/font-size:\s*([^;]+)/i)
          if (fontSizeMatch) {
            let fontSize = fontSizeMatch[1].trim()
            if (fontSize.includes('pt')) {
              const ptValue = parseFloat(fontSize.replace('pt', ''))
              fontSize = Math.round(ptValue * 1.33) + 'px'
            }
            cellStyle.fontsize = fontSize
          }
          
          const fontFamilyMatch = inlineStyle.match(/font-family:\s*([^;]+)/i)
          if (fontFamilyMatch) {
            let fontFamily = fontFamilyMatch[1].trim().replace(/['"]/g, '')
            const fontMap: Record<string, string> = {
              '黑体': 'SimHei',
              '微软雅黑': 'Microsoft YaHei',
              '宋体': 'SimSun',
              '楷体': 'KaiTi'
            }
            cellStyle.fontname = fontMap[fontFamily] || fontFamily
          }
          
          const fontWeightMatch = inlineStyle.match(/font-weight:\s*([^;]+)/i)
          if (fontWeightMatch) {
            const fontWeight = fontWeightMatch[1].trim()
            cellStyle.bold = fontWeight === 'bold' || parseInt(fontWeight) >= 600
          }
          
          const fontStyleMatch = inlineStyle.match(/font-style:\s*([^;]+)/i)
          if (fontStyleMatch && fontStyleMatch[1].trim() === 'italic') {
            cellStyle.em = true
          }
          
          const textDecorationMatch = inlineStyle.match(/text-decoration:\s*([^;]+)/i)
          if (textDecorationMatch) {
            const decoration = textDecorationMatch[1].trim()
            if (decoration.includes('underline')) cellStyle.underline = true
            if (decoration.includes('line-through')) cellStyle.strikethrough = true
          }

          // 解析文本对齐
          const textAlignMatch = inlineStyle.match(/text-align:\s*([^;]+)/i)
          if (textAlignMatch) {
            cellStyle.align = textAlignMatch[1].trim() as TextAlign
          } else {
            // 如果没有明确的对齐设置，根据列位置设置默认对齐方式
            // 第一列默认左对齐，其他列默认居中对齐
            cellStyle.align = colIndex === 0 ? 'left' : 'center'
          }

          // 创建真实的单元格对象
          const realCell: TableCell = {
            id: nanoid(10),
            colspan,
            rowspan,
            text: cellText,
            style: Object.keys(cellStyle).length > 0 ? cellStyle : undefined
          }

          // 在占用矩阵中填充单元格
          for (let r = rowIndex; r < rowIndex + rowspan; r++) {
            // 确保有足够的行
            while (occupiedMatrix.length <= r) {
              occupiedMatrix.push([])
            }
            for (let c = colIndex; c < colIndex + colspan; c++) {
              if (r === rowIndex && c === colIndex) {
                // 原点位置放置真实单元格
                occupiedMatrix[r][c] = realCell
              } else {
                // 被合并的位置放置空单元格，也要设置对齐方式
                occupiedMatrix[r][c] = {
                  id: nanoid(10),
                  colspan: 1,
                  rowspan: 1,
                  text: '',
                  style: {
                    align: c === 0 ? 'left' : 'center'
                  }
                }
              }
            }
          }

          maxCols = Math.max(maxCols, colIndex + colspan)
        })
      })

      // 第二遍：确保所有行都有相同的列数，并填充缺失的位置
      const tableData: TableCell[][] = []
      
      for (let r = 0; r < occupiedMatrix.length; r++) {
        const row: TableCell[] = []
        
        for (let c = 0; c < maxCols; c++) {
          if (occupiedMatrix[r][c] !== undefined) {
            row.push(occupiedMatrix[r][c]!)
          } else {
            // 如果某个位置没有被填充，创建一个空单元格
            row.push({
              id: nanoid(10),
              colspan: 1,
              rowspan: 1,
              text: '',
              style: {
                align: c === 0 ? 'left' : 'center' // 第一列左对齐，其他列居中
              }
            })
          }
        }
        
        tableData.push(row)
      }

      // 计算列宽
      const colWidths: number[] = Array(maxCols).fill(1 / maxCols)

      // 计算表格尺寸
      const estimatedWidth = Math.max(500, Math.min(900, maxCols * 100))
      const estimatedHeight = Math.max(tableData.length * defaultRowHeight, rows.length * defaultRowHeight)

      // 检测主题色和表头
      let themeColor = 'rgb(255, 255, 255)' // 默认主题色
      let hasRowHeader = false
      let hasColHeader = false
      
      // 首先尝试从table的样式中获取主题色
      const tableStyle = table.getAttribute('style') || ''
      const themeColorMatch = tableStyle.match(/--themeColor:\s*([^;)]+)/i)
      if (themeColorMatch) {
        themeColor = cleanColorValue(themeColorMatch[1])
      }
      
      // 检测表头类型
      if (tableData.length > 0) {
        // 检查第一行是否为列头（检查有内容的单元格）
        const firstRow = tableData[0]
        if (firstRow && firstRow.length > 0) {
          const contentCells = firstRow.filter(cell => cell.text.trim() !== '')
          hasColHeader = contentCells.some(cell => 
            cell.style?.backcolor || 
            cell.style?.bold || 
            (cell.style?.color && cell.style.color !== 'rgb(51, 51, 51)')
          )
        }
        
        // 检查第一列是否为行头（检查每行第一个有内容的单元格）
        if (tableData.length > 1) {
          const firstColCells = tableData.map(row => row[0]).filter(cell => cell && cell.text.trim() !== '')
          if (firstColCells.length > 1) {
            hasRowHeader = firstColCells.some(cell => 
              cell.style?.backcolor || 
              cell.style?.bold || 
              (cell.style?.color && cell.style?.color !== 'rgb(51, 51, 51)')
            )
          }
        }
      }

      // 获取边框样式
      const getOutlineStyle = () => {
        if (rows.length > 0) {
          const firstCell = rows[0].querySelector('td, th')
          if (firstCell) {
            const style = firstCell.getAttribute('style') || ''
            const borderWidthMatch = style.match(/border-width:\s*([^;]+)/i)
            const borderColorMatch = style.match(/border-color:\s*([^;]+)/i)
            const borderStyleMatch = style.match(/border-style:\s*([^;]+)/i)
            
            return {
              width: borderWidthMatch ? parseInt(borderWidthMatch[1]) : 1,
              style: borderStyleMatch ? borderStyleMatch[1].trim() as any : 'solid',
              color: borderColorMatch ? cleanColorValue(borderColorMatch[1]) : '#cccccc'
            }
          }
        }
        return { width: 1, style: 'solid' as const, color: '#cccccc' }
      }

      const tableElement: PPTTableElement = {
        type: 'table',
        id: nanoid(10),
        width: estimatedWidth,
        height: estimatedHeight,
        colWidths,
        rotate: 0,
        data: tableData,
        left: 100,
        top: 150,
        outline: getOutlineStyle(),
        theme: {
          color: themeColor,
          rowHeader: hasRowHeader,
          rowFooter: false,
          colHeader: hasColHeader,
          colFooter: false
        },
        cellMinHeight: defaultRowHeight
      }

      console.log('=== Table Structure ===')
      console.log(`Rows: \${tableData.length}, Cols: \${maxCols}`)
      tableData.forEach((row, i) => {
        console.log(`Row \${i} (\${row.length} cells):`, row.map(cell => ({
          text: cell.text || '(empty)',
          colspan: cell.colspan,
          rowspan: cell.rowspan,
          align: cell.style?.align || 'default'
        })))
      })
      
      return tableElement

    } catch (error) {
      console.error('Error parsing HTML table:', error)
      return null
    }
  }

  const AIPPT = (templateSlides: Slide[], _AISlides: AIPPTSlide[], imgs?: ImgPoolItem[]) => {
    slidesStore.updateSlideIndex(slidesStore.slides.length - 1)

    if (imgs) imgPool.value = imgs

    const AISlides: AIPPTSlide[] = []
    for (const template of _AISlides) {
      if (template.type === 'content') {
        const items = template.data.items
        if (items.length === 5 || items.length === 6) {
          const items1 = items.slice(0, 3)
          const items2 = items.slice(3)
          AISlides.push({ ...template, data: { ...template.data, items: items1 } })
          AISlides.push({ ...template, data: { ...template.data, items: items2 }, offset: 3 })
        }
        else if (items.length === 7 || items.length === 8) {
          const items1 = items.slice(0, 4)
          const items2 = items.slice(4)
          AISlides.push({ ...template, data: { ...template.data, items: items1 } })
          AISlides.push({ ...template, data: { ...template.data, items: items2 }, offset: 4 })
        }
        else if (items.length === 9 || items.length === 10) {
          const items1 = items.slice(0, 3)
          const items2 = items.slice(3, 6)
          const items3 = items.slice(6)
          AISlides.push({ ...template, data: { ...template.data, items: items1 } })
          AISlides.push({ ...template, data: { ...template.data, items: items2 }, offset: 3 })
          AISlides.push({ ...template, data: { ...template.data, items: items3 }, offset: 6 })
        }
        else if (items.length > 10) {
          const items1 = items.slice(0, 4)
          const items2 = items.slice(4, 8)
          const items3 = items.slice(8)
          AISlides.push({ ...template, data: { ...template.data, items: items1 } })
          AISlides.push({ ...template, data: { ...template.data, items: items2 }, offset: 4 })
          AISlides.push({ ...template, data: { ...template.data, items: items3 }, offset: 8 })
        }
        else {
          AISlides.push(template)
        }
      }
      else if (template.type === 'contents') {
        const items = template.data.items
        if (items.length === 11) {
          const items1 = items.slice(0, 6)
          const items2 = items.slice(6)
          AISlides.push({ ...template, data: { ...template.data, items: items1 } })
          AISlides.push({ ...template, data: { ...template.data, items: items2 }, offset: 6 })
        }
        else if (items.length > 11) {
          const items1 = items.slice(0, 10)
          const items2 = items.slice(10)
          AISlides.push({ ...template, data: { ...template.data, items: items1 } })
          AISlides.push({ ...template, data: { ...template.data, items: items2 }, offset: 10 })
        }
        else {
          AISlides.push(template)
        }
      }
      else AISlides.push(template)
    }

    const coverTemplates = templateSlides.filter(slide => slide.type === 'cover')
    const contentsTemplates = templateSlides.filter(slide => slide.type === 'contents')
    const transitionTemplates = templateSlides.filter(slide => slide.type === 'transition')
    const contentTemplates = templateSlides.filter(slide => slide.type === 'content')
    const endTemplates = templateSlides.filter(slide => slide.type === 'end')

    if (!transitionTemplate.value) {
      const _transitionTemplate = transitionTemplates[Math.floor(Math.random() * transitionTemplates.length)]
      transitionTemplate.value = _transitionTemplate
    }

    const slides = []
    
    for (const item of AISlides) {
      if (item.type === 'cover') {
        const coverTemplate = coverTemplates[Math.floor(Math.random() * coverTemplates.length)]
        const elements = coverTemplate.elements.map(el => {
          if (el.type === 'image' && el.imageType && imgPool.value.length) return getNewImgElement(el)
          if (el.type !== 'text' && el.type !== 'shape') return el
          if (checkTextType(el, 'title') && item.data.title) {
            return getNewTextElement({ el, text: item.data.title, maxLine: 1 })
          }
          if (checkTextType(el, 'content') && item.data.text) {
            return getNewTextElement({ el, text: item.data.text, maxLine: 3 })
          }
          return el
        })
        slides.push({
          ...coverTemplate,
          id: nanoid(10),
          elements,
        })
      }
      else if (item.type === 'contents') {
        const _contentsTemplates = getUseableTemplates(contentsTemplates, item.data.items.length, 'item')
        const contentsTemplate = _contentsTemplates[Math.floor(Math.random() * _contentsTemplates.length)]

        const sortedNumberItems = contentsTemplate.elements.filter(el => checkTextType(el, 'itemNumber'))
        const sortedNumberItemIds = sortedNumberItems.sort((a, b) => {
          if (sortedNumberItems.length > 6) {
            let aContent = ''
            let bContent = ''
            if (a.type === 'text') aContent = a.content
            if (a.type === 'shape') aContent = a.text!.content
            if (b.type === 'text') bContent = b.content
            if (b.type === 'shape') bContent = b.text!.content

            if (aContent && bContent) {
              const aIndex = parseInt(aContent)
              const bIndex = parseInt(bContent)

              return aIndex - bIndex
            }
          }
          const aIndex = a.left + a.top * 2
          const bIndex = b.left + b.top * 2
          return aIndex - bIndex
        }).map(el => el.id)

        const sortedItems = contentsTemplate.elements.filter(el => checkTextType(el, 'item'))
        const sortedItemIds = sortedItems.sort((a, b) => {
          if (sortedItems.length > 6) {
            const aItemNumber = sortedNumberItems.find(item => item.groupId === a.groupId)
            const bItemNumber = sortedNumberItems.find(item => item.groupId === b.groupId)

            if (aItemNumber && bItemNumber) {
              let aContent = ''
              let bContent = ''
              if (aItemNumber.type === 'text') aContent = aItemNumber.content
              if (aItemNumber.type === 'shape') aContent = aItemNumber.text!.content
              if (bItemNumber.type === 'text') bContent = bItemNumber.content
              if (bItemNumber.type === 'shape') bContent = bItemNumber.text!.content
  
              if (aContent && bContent) {
                const aIndex = parseInt(aContent)
                const bIndex = parseInt(bContent)
  
                return aIndex - bIndex
              }
            }
          }

          const aIndex = a.left + a.top * 2
          const bIndex = b.left + b.top * 2
          return aIndex - bIndex
        }).map(el => el.id)

        const longestText = item.data.items.reduce((longest, current) => current.length > longest.length ? current : longest, '')

        const unusedElIds: string[] = []
        const unusedGroupIds: string[] = []
        const elements = contentsTemplate.elements.map(el => {
          if (el.type === 'image' && el.imageType && imgPool.value.length) return getNewImgElement(el)
          if (el.type !== 'text' && el.type !== 'shape') return el
          if (checkTextType(el, 'item')) {
            const index = sortedItemIds.findIndex(id => id === el.id)
            const itemTitle = item.data.items[index]
            if (itemTitle) return getNewTextElement({ el, text: itemTitle, maxLine: 1, longestText })

            unusedElIds.push(el.id)
            if (el.groupId) unusedGroupIds.push(el.groupId)
          }
          if (checkTextType(el, 'itemNumber')) {
            const index = sortedNumberItemIds.findIndex(id => id === el.id)
            const offset = item.offset || 0
            return getNewTextElement({ el, text: index + offset + 1 + '', maxLine: 1, digitPadding: true })
          }
          return el
        }).filter(el => !unusedElIds.includes(el.id) && !(el.groupId && unusedGroupIds.includes(el.groupId)))
        slides.push({
          ...contentsTemplate,
          id: nanoid(10),
          elements,
        })
      }
      else if (item.type === 'transition') {
        transitionIndex.value = transitionIndex.value + 1
        const elements = transitionTemplate.value.elements.map(el => {
          if (el.type === 'image' && el.imageType && imgPool.value.length) return getNewImgElement(el)
          if (el.type !== 'text' && el.type !== 'shape') return el
          if (checkTextType(el, 'title') && item.data.title) {
            return getNewTextElement({ el, text: item.data.title, maxLine: 1 })
          }
          if (checkTextType(el, 'content') && item.data.text) {
            return getNewTextElement({ el, text: item.data.text, maxLine: 3 })
          }
          if (checkTextType(el, 'partNumber')) {
            return getNewTextElement({ el, text: transitionIndex.value + '', maxLine: 1, digitPadding: true })
          }
          return el
        })
        slides.push({
          ...transitionTemplate.value,
          id: nanoid(10),
          elements,
        })
      }
      else if (item.type === 'content') {
        // 检查是否包含图片内容
        const hasImages = item.data.items.some(item => item.image)
        const imageItems = item.data.items.filter(item => item.image)
        
        // 检查是否包含表格内容（在items中）
        const hasTable = item.data.items.some(item => item.table)
        const tableItems = item.data.items.filter(item => item.table)
        
        let _contentTemplates = getUseableTemplates(contentTemplates, item.data.items.length, 'item')
        
        // 如果有图片内容，优先选择包含图片元素的模板
        if (hasImages) {
          const templatesWithImages = _contentTemplates.filter(template => 
            template.elements.some(el => el.type === 'image' && el.imageType === 'itemFigure')
          )
          if (templatesWithImages.length > 0) {
            _contentTemplates = templatesWithImages
          }
        }
        
        const contentTemplate = _contentTemplates[Math.floor(Math.random() * _contentTemplates.length)]

        const sortedTitleItemIds = contentTemplate.elements.filter(el => checkTextType(el, 'itemTitle')).sort((a, b) => {
          const aIndex = a.left + a.top * 2
          const bIndex = b.left + b.top * 2
          return aIndex - bIndex
        }).map(el => el.id)

        const sortedTextItemIds = contentTemplate.elements.filter(el => checkTextType(el, 'item')).sort((a, b) => {
          const aIndex = a.left + a.top * 2
          const bIndex = b.left + b.top * 2
          return aIndex - bIndex
        }).map(el => el.id)

        const sortedNumberItemIds = contentTemplate.elements.filter(el => checkTextType(el, 'itemNumber')).sort((a, b) => {
          const aIndex = a.left + a.top * 2
          const bIndex = b.left + b.top * 2
          return aIndex - bIndex
        }).map(el => el.id)

        // 获取模板中的图片元素并排序
        const sortedImageItemIds = contentTemplate.elements.filter(el => 
          el.type === 'image' && el.imageType === 'itemFigure'
        ).sort((a, b) => {
          const aIndex = a.left + a.top * 2
          const bIndex = b.left + b.top * 2
          return aIndex - bIndex
        }).map(el => el.id)

        const itemTitles = []
        const itemTexts = []

        for (const _item of item.data.items) {
          if (_item.title) itemTitles.push(_item.title)
          if (_item.text) itemTexts.push(_item.text)
        }
        const longestTitle = itemTitles.reduce((longest, current) => current.length > longest.length ? current : longest, '')
        const longestText = itemTexts.reduce((longest, current) => current.length > longest.length ? current : longest, '')

        // 创建新的图片元素来替换现有图片或添加新图片
        const createAIImageElement = (templateEl: PPTImageElement, imageSrc: string): PPTImageElement => {
          return {
            ...templateEl,
            id: nanoid(10),
            src: imageSrc,
            // 移除裁剪设置，显示完整图片
            clip: undefined,
            // 移除滤镜效果
            filters: undefined,
          }
        }

        // 如果内容只有图片，需要创建一个以图片为主的布局
        const isImageOnlyContent = item.data.items.length === 1 && item.data.items[0].image && 
                                  !item.data.items[0].title && !item.data.items[0].text

        let unusedElIds: string[] = []
        let unusedGroupIds: string[] = []

        let elements = contentTemplate.elements.map(el => {
          // 处理模板中的装饰性图片
          if (el.type === 'image' && el.imageType === 'pageFigure' && imgPool.value.length) {
            return getNewImgElement(el)
          }
          
          // 处理模板中的项目图片
          if (el.type === 'image' && el.imageType === 'itemFigure') {
            const index = sortedImageItemIds.findIndex(id => id === el.id)
            const contentItem = item.data.items[index]
            if (contentItem && contentItem.image) {
              return createAIImageElement(el, contentItem.image)
            }
            // 如果没有对应的AI图片，但该位置的item有表格，则移除这个图片元素
            if (contentItem && contentItem.table) {
              unusedElIds.push(el.id)
              if (el.groupId) unusedGroupIds.push(el.groupId)
              return el // 先返回，后面会被过滤掉
            }
            // 如果没有对应的AI图片，保持原图片或使用图片池中的图片
            return imgPool.value.length ? getNewImgElement(el) : el
          }
          
          if (el.type !== 'text' && el.type !== 'shape') return el
          
          if (item.data.items.length === 1) {
            const contentItem = item.data.items[0]
            if (checkTextType(el, 'content') && contentItem.text) {
              return getNewTextElement({ el, text: contentItem.text, maxLine: 6 })
            }
          }
          else {
            if (checkTextType(el, 'itemTitle')) {
              const index = sortedTitleItemIds.findIndex(id => id === el.id)
              const contentItem = item.data.items[index]
              if (contentItem && contentItem.title) {
                return getNewTextElement({ el, text: contentItem.title, longestText: longestTitle, maxLine: 1 })
              }
              // 如果没有对应的标题，移除这个元素
              unusedElIds.push(el.id)
              if (el.groupId) unusedGroupIds.push(el.groupId)
            }
            if (checkTextType(el, 'item')) {
              const index = sortedTextItemIds.findIndex(id => id === el.id)
              const contentItem = item.data.items[index]
              if (contentItem && contentItem.text) {
                return getNewTextElement({ el, text: contentItem.text, longestText, maxLine: 4 })
              }
              // 如果没有对应的文本，移除这个元素
              unusedElIds.push(el.id)
              if (el.groupId) unusedGroupIds.push(el.groupId)
            }
            if (checkTextType(el, 'itemNumber')) {
              const index = sortedNumberItemIds.findIndex(id => id === el.id)
              const contentItem = item.data.items[index]
              const offset = item.offset || 0
              if (contentItem) {
                return getNewTextElement({ el, text: index + offset + 1 + '', maxLine: 1, digitPadding: true })
              }
              // 如果没有对应的项目，移除这个元素
              unusedElIds.push(el.id)
              if (el.groupId) unusedGroupIds.push(el.groupId)
            }
          }
          if (checkTextType(el, 'title') && item.data.title) {
            return getNewTextElement({ el, text: item.data.title, maxLine: 1 })
          }
          return el
        }).filter(el => !unusedElIds.includes(el.id) && !(el.groupId && unusedGroupIds.includes(el.groupId)))

        // 处理表格（在items中）
        if (hasTable && tableItems.length > 0) {
          // 计算已有内容的边界，用于合理布局表格
          const titleElements = elements.filter(el => checkTextType(el, 'title'))
          const textElements = elements.filter(el => checkTextType(el, 'item') || checkTextType(el, 'itemTitle'))
          
          let currentY = 150
          if (titleElements.length > 0) {
            const maxTitleBottom = Math.max(...titleElements.map(el => el.top + el.height))
            currentY = maxTitleBottom + 20
          }

          // 为每个包含表格的item创建表格元素
          tableItems.forEach((tableItem, tableIndex) => {
            if (tableItem.table) {
              const tableElement = parseHTMLTable(tableItem.table)
              if (tableElement) {
                // 找到对应的文本元素位置，将表格放在其下方
                const itemIndex = item.data.items.findIndex(dataItem => dataItem === tableItem)
                
                // 计算表格位置
                let tableTop = currentY
                let tableLeft = 100
                
                // 如果有对应的文本元素，根据其位置调整表格位置
                if (textElements[itemIndex]) {
                  const textEl = textElements[itemIndex]
                  tableTop = textEl.top + textEl.height + 10
                  tableLeft = textEl.left
                  
                  // 确保表格不会超出幻灯片边界
                  const slideWidth = 1000
                  if (tableLeft + tableElement.width > slideWidth - 50) {
                    tableLeft = slideWidth - tableElement.width - 50
                  }
                } else {
                  // 如果没有对应的文本元素，使用网格布局
                  const slideWidth = 1000
                  const tablesPerRow = Math.min(2, tableItems.length) // 每行最多2个表格
                  const tableWidth = (slideWidth - 100) / tablesPerRow - 20
                  
                  tableElement.width = Math.min(tableElement.width, tableWidth)
                  
                  const row = Math.floor(tableIndex / tablesPerRow)
                  const col = tableIndex % tablesPerRow
                  
                  tableLeft = 50 + col * (tableWidth + 20)
                  tableTop = currentY + row * (tableElement.height + 30)
                }
                
                tableElement.left = tableLeft
                tableElement.top = tableTop
                
                elements.push(tableElement)
                
                // 更新当前Y位置，为下一个元素预留空间
                currentY = Math.max(currentY, tableTop + tableElement.height + 20)
              }
            }
          })
        }

        // 如果是纯图片内容且模板中没有合适的图片位置，创建居中的大图片
        if (isImageOnlyContent && !sortedImageItemIds.length && item.data.items[0].image) {
          const imageElement: PPTImageElement = {
            type: 'image',
            id: nanoid(10),
            src: item.data.items[0].image,
            width: 600,
            height: 400,
            left: 200,
            top: 100,
            fixedRatio: true,
            rotate: 0,
            imageType: 'itemFigure'
          }
          elements.push(imageElement)
        }
        // 如果有多个图片但模板图片位置不够，添加额外的图片元素
        else if (imageItems.length > sortedImageItemIds.length) {
          const extraImages = imageItems.slice(sortedImageItemIds.length)
          extraImages.forEach((imgItem, index) => {
            if (imgItem.image) {
              const imageElement: PPTImageElement = {
                type: 'image',
                id: nanoid(10),
                src: imgItem.image,
                width: 300,
                height: 200,
                left: 50 + (index * 350),
                top: 350,
                fixedRatio: true,
                rotate: 0,
                imageType: 'itemFigure'
              }
              elements.push(imageElement)
            }
          })
        }

        slides.push({
          ...contentTemplate,
          id: nanoid(10),
          elements,
        })
      }
      else if (item.type === 'end') {
        const endTemplate = endTemplates[Math.floor(Math.random() * endTemplates.length)]
        const elements = endTemplate.elements.map(el => {
          if (el.type === 'image' && el.imageType && imgPool.value.length) return getNewImgElement(el)
          return el
        })
        slides.push({
          ...endTemplate,
          id: nanoid(10),
          elements,
        })
      }
    }
    if (isEmptySlide.value) slidesStore.setSlides(slides)
    else addSlidesFromData(slides)
  }

  return {
    presetImgPool,
    AIPPT,
    getMdContent,
    getJSONContent,
  }
}