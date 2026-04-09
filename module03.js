function setDatas(){
	const baseSettingAndDatas = getDatas()
	if (!baseSettingAndDatas){return}
    const dayAverage = toVal(ThisWorkbook.Worksheets.Item("操作").Range("D2").Value2)
    const start_day = Application.WorksheetFunction.Text(ThisWorkbook.Worksheets.Item("操作").Range("D3").Value2,"yyyy-MM-dd")
    const end_day = Application.WorksheetFunction.Text(ThisWorkbook.Worksheets.Item("操作").Range("D4").Value2,"yyyy-MM-dd")  
    if (!start_day || !end_day){
        MsgBox("请先设置开始日期和结束日期")
        return
    }
    Application.DisplayAlerts = false
    var wst = ThisWorkbook.Worksheets.Item("模板")
    try{
        wst.Copy()
        var newWb = Application.ActiveWorkbook
        wst = newWb.Worksheets.Item("模板")
        var level = setAloneCategory(wst,'公司','商品库存统计',baseSettingAndDatas)
        var temp01 = setAloneShop(wst,'公司',level,baseSettingAndDatas,dayAverage,"D")
        var temp02 = setAloneShop(wst,'星耀店',level,baseSettingAndDatas,dayAverage,"R")
        var temp03 = setAloneShop(wst,'八中店',level,baseSettingAndDatas,dayAverage,"AF")
        var temp04 = setAloneShop(wst,'五中店',level,baseSettingAndDatas,dayAverage,"AT")
        var temp05 = setAloneShop(wst,'十四中店',level,baseSettingAndDatas,dayAverage,"BH")
        
        var lastRow = level.length + 1
        var arr = wst.Range(`B2:D${lastRow-1}`).Value2
        wst.Range(`B2:D${lastRow-1}`).Value2 = sortByColumn3(arr)

        wst.Rows(1).Copy()
        wst.Rows(1).PasteSpecial(xlPasteValues, xlPasteSpecialOperationNone, false, false)
        Application.CutCopyMode = false
        wst.Rows(2).Copy()
        wst.Rows("3:" + lastRow).PasteSpecial(xlPasteFormats)
        Application.CutCopyMode = false
        wst.Range("D2").Select()
        ActiveWindow.FreezePanes = true;
		const savePath = `${ThisWorkbook.Path}\\二级分类动销表（月至今销量）${start_day}-${end_day}.xlsx`
        newWb.SaveAs(savePath)
        newWb.Close()
        MsgBox(`运行完成，保存位置：${savePath}`)
    }catch(e){
        try{
            newWb.Close()
        }catch(e2){
            
        }
        MsgBox(`设置数据失败：${e.message}`)
    }

    Application.DisplayAlerts = true
}

function setAloneShop(wst,shopName,level,baseSettingAndDatas,dayAverage,writeCol){
    const shop_info = baseSettingAndDatas[shopName]
    var result = []
    var start_index = 0
    for (let i=0;i<level.length;i++){
        const category_name = level[i][2]
        var temp_row = []
        if (level[i][1] === ""){
            start_index = start_index +1        
            temp_row.push(safeGet(shop_info,"营业占比分析下","datas",category_name,"销售额"))
            temp_row.push(safeGet(shop_info,"营业占比分析上","datas",category_name,"销售额"))
            temp_row.push(`=IFERROR(RC[-2]-RC[-1],0)`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(safeGet(shop_info,"营业占比分析下","datas",category_name,"商品数量"))
            temp_row.push(safeGet(shop_info,"营业占比分析上","datas",category_name,"商品数量"))
            temp_row.push(`=IFERROR(RC[-2]-RC[-1],0)`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(safeGet(shop_info,"营业占比分析下","datas",category_name,"销售成本"))
            temp_row.push(`=IFERROR(RC[-1]/${dayAverage},0)`)
            temp_row.push(safeGet(shop_info,"商品库存统计","datas",category_name,"成本总额"))
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(`=INDEX(R[1]C:R[1048]C,MATCH("汇总",R[1]C[-${toCol(writeCol)+10}]:R[1048]C[-${toCol(writeCol)+10}],0))`)
            temp_row.push(`=IFERROR((RC[-1]-RC[-2])*RC[-4],0)`)
        }else if(level[i][1] === "汇总"){
            temp_row.push(`=SUM(R[${-start_index}]C:R[-1]C)`)
            temp_row.push(`=SUM(R[${-start_index}]C:R[-1]C)`)
            temp_row.push(`=IFERROR(RC[-2]-RC[-1],0)`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(`=SUM(R[${-start_index}]C:R[-1]C)`)
            temp_row.push(`=SUM(R[${-start_index}]C:R[-1]C)`)
            temp_row.push(`=IFERROR(RC[-2]-RC[-1],0)`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(`=SUM(R[${-start_index}]C:R[-1]C)`)
            temp_row.push(`=IFERROR(RC[-1]/${dayAverage},0)`)
            temp_row.push(`=SUM(R[${-start_index}]C:R[-1]C)`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(safeGet(shop_info,"商品库存统计","Level01",category_name,"targetDay"))
            temp_row.push(`=IFERROR((RC[-1]-RC[-2])*RC[-4],0)`)
            start_index = 0
        }else if(level[i][1] === "合计"){
            temp_row.push(`=SUMIFS(R2C:R[-1]C,R2C2:R[-1]C2,"汇总")`)
            temp_row.push(`=SUMIFS(R2C:R[-1]C,R2C2:R[-1]C2,"汇总")`)
            temp_row.push(`=IFERROR(RC[-2]-RC[-1],0)`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(`=SUMIFS(R2C:R[-1]C,R2C2:R[-1]C2,"汇总")`)
            temp_row.push(`=SUMIFS(R2C:R[-1]C,R2C2:R[-1]C2,"汇总")`)
            temp_row.push(`=IFERROR(RC[-2]-RC[-1],0)`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(`=SUMIFS(R2C:R[-1]C,R2C2:R[-1]C2,"汇总")`)
            temp_row.push(`=IFERROR(RC[-1]/${dayAverage},0)`)
            temp_row.push(`=SUMIFS(R2C:R[-1]C,R2C2:R[-1]C2,"汇总")`)
            temp_row.push(`=IFERROR(RC[-1]/RC[-2],0)`)
            temp_row.push(`=SUMIFS(R2C:R[-1]C,R2C2:R[-1]C2,"汇总")`)
            temp_row.push(`=IFERROR((RC[-1]-RC[-2])*RC[-4],0)`)
        }
        
        result.push(temp_row)
    }
    wst.Range(`${writeCol}2`).Resize(result.length,result[0].length).FormulaR1C1 = result
    return result
}
/*
 * 单独处理商品库存统计
 * level01 的格式是 {“A1-本册”：{"key": ["A1","","本册"]}, ...}
 * level02 的格式是 {“A1-1横线B5/A4”：{"key": ["A1",1,"横线B5/A4"]}, ...}
 */
function setAloneCategory(wst,shopName,category_name,baseSettingAndDatas){
    const level01 = baseSettingAndDatas[shopName][category_name].Level01
    const level02 = baseSettingAndDatas[shopName][category_name].datas

    const level02Grouped = {}
    for (let name in level02){
        const key0 = level02[name].key[0]
        if (!level02Grouped[key0]){
            level02Grouped[key0] = []
        }
        level02Grouped[key0].push({ name: name, data: level02[name] })
    }
    for (let key0 in level02Grouped){
        level02Grouped[key0].sort((a, b) => {
            const diff = a.data.key[1] - b.data.key[1]
            if (diff !== 0) return diff
            if (a.data.key[2] < b.data.key[2]) return -1
            if (a.data.key[2] > b.data.key[2]) return 1
            return 0
        })
    }

    const level01Sorted = Object.keys(level01).sort((a, b) => {
        return level01[a].key[0].localeCompare(level01[b].key[0])
    })

    const result = []
    let seq = 1
    for (let name of level01Sorted){
        const key0 = level01[name].key[0]
        if (level02Grouped[key0]){
            for (let item of level02Grouped[key0]){
                result.push([seq, "", item.name])
                seq++
            }
            level01[name].datas = {}
            for (let item of level02Grouped[key0]){
                level01[name].datas[item.name] = item.data
            }
        }
        result.push([seq, "汇总", name])
        seq++
    }
    result.push(["", "合计", ""])
    wst.Range("A2").Resize(result.length,3).Value2 = result
    return result
}




function getDatas(){
	var baseSetting = getBasedSetting()
	if (!baseSetting){
		MsgBox("配置参数错误，请注意检查！")
		return null
	}
	Object.keys(baseSetting).forEach((shop_name,i) =>{
		const shop_info = baseSetting[shop_name]
		Object.keys(shop_info).forEach((category_name,j) =>{
			DoEvents();
			category = shop_info[category_name]
			if (!category.path){
				MsgBox(`${shop_name} 选中的路径不存在，请注意检查配置！然后再运行...`)
				return null
			}	
			if (category_name === "商品库存统计"){
				category.datas = aopenFileAndRead(shop_name,category_name,category)
			}else if(category_name === "营业占比分析上"){
				category.datas = aopenFileAndRead(shop_name,category_name,category)					
			}else if(category_name === "营业占比分析下"){
				category.datas = aopenFileAndRead(shop_name,category_name,category)					
			}
			if (!category.datas){
				MsgBox(`${shop_info} - ${category_name} - 数据获取异常，请注意检查！`)
				return null
			}
		})
	})
	
	return baseSetting
}

function aopenFileAndRead(shopName,categoryName,category){
	Application.DisplayAlerts = false
	var result = {}
	try{
		const path = category.subPaths[Object.keys(category.subPaths)[0]].path
		var wb = Application.Workbooks.Open(FileName=path,ReadOnly=true);
		var wst = wb.Worksheets.Item(1);
		const usedRow = wst.Cells(wst.UsedRange.Rows.Count+256,"A").End(xlUp).Row
        const usedCol = wst.Cells(wst.UsedRange.Rows.Count+256,"A").End(xlToRight).Column
		if (usedRow <= 1){
			return null
		}
		const arr = wst.Range("A1").Resize(usedRow,usedCol).Value2
        var tempClass = ""
		for (let i = 1; i < usedRow;i++ ){
            tempClass = '商品分类'
            const shop_class_index = category.fields[tempClass].columnIndex
            if (!shop_class_index){
                MsgBox(`${shopName} - ${categoryName} - ${tempClass} 字段不存在，请注意检查！`)
                return null
            }
			var tempStr = arr[i][shop_class_index-1]
            if (tempStr === "" || !tempStr){ continue }
            if (tempStr === "总计" || tempStr === "汇总" || tempStr === "合计"){ continue }
            result[tempStr] = {}
            Object.keys(category.fields).forEach((field_name) =>{
                if (field_name === tempClass){ 
                	result[tempStr]["key"] = splitCategory(arr[i][shop_class_index-1])
                	
                }else{
	                var temp_index = category.fields[field_name].columnIndex
	                if (!temp_index){
	                    MsgBox(`${shopName} - ${categoryName} - ${tempClass} 字段不存在，请注意检查！`)
	                    return null
	                }
	                result[tempStr][field_name] = toVal(arr[i][temp_index-1])                	
                }
            })
			DoEvents();
		}
        wb.Close()
	}catch(error){
		try{
        	wb.Close()			
		}catch(error){
		}
		Application.DisplayAlerts = true
		return null
	}
	DoEvents();
	return result
}


/*
 * @return 格式：{
 * 	shop_name: { 
 * 		"商品库存统计": { 
 * 			path:"",
 *  		year:"",
 *  		fields:{ 
 * 				"分类-字段名": {"column":"B-字段位置","columnIndex":2},
 * 				...
 *  		}
 *  	} , 
 *  	"营业占比分析上": ... 
 * }
*/
function getBasedSetting(){
	const wst = ThisWorkbook.Worksheets("操作")
	var result = {}
	var temp01 = _getAloneInfo("2","F")
	if (!temp01){ return null }
	var temp02 = _getAloneInfo("3","I")
	if (!temp02){ return null }
	var temp03 = _getAloneInfo("4","L")
	if (!temp03){ return null }
    //一级分类
    const category_level01_area = "U" 
    var category_temp04 = {}
	var usedRow = wst.Cells(wst.UsedRange.Rows.Count+256,category_level01_area).End(xlUp).Row
    if (usedRow <= 1){
		MsgBox(`一级分类 没有配置，请注意检查配置！`)
		return null
	}
    const category_arr = wst.Range(`${category_level01_area}1`).Resize(usedRow,2).Value2
    for(let r=1; r< usedRow;r++){
        tempStr = category_arr[r][0]
        if (tempStr === ''){ continue}
        category_temp04[tempStr] = {
            key: splitCategory(tempStr),
            targetDay: toVal(category_arr[r][1]),     
            datas: {}
        }
    }
    if (!category_temp04){
		MsgBox(`一级分类 没有配置，请注意检查配置！`)
		return null
    }

	//固定店铺
	const shopArea = "R"
	usedRow = wst.Cells(wst.UsedRange.Rows.Count+256,shopArea).End(xlUp).Row
	if (usedRow <= 1){
		MsgBox(`固定店铺 没有配置，请注意检查配置！`)
		return null
	}
	const arr = wst.Range(`${shopArea}1`).Resize(usedRow,1).Value2
	for (let i = 1; i < usedRow;i++ ){
		var tempStr = arr[i][0]
		if (tempStr === "" || !tempStr){ continue }
		var _temp01 = JSON.parse(JSON.stringify(temp01))
		_temp01.subPaths = getFilesByPath(temp01.path,[tempStr],["~$"],["xlsx","xls","xlsm"])
        _temp01.Level01 = JSON.parse(JSON.stringify(category_temp04))
		var _temp02 = JSON.parse(JSON.stringify(temp02))
		_temp02.subPaths = getFilesByPath(temp02.path,[tempStr],["~$"],["xlsx","xls","xlsm"])
        _temp02.Level01 = JSON.parse(JSON.stringify(category_temp04))
		var _temp03 = JSON.parse(JSON.stringify(temp03))
		_temp03.subPaths = getFilesByPath(temp03.path,[tempStr],["~$"],["xlsx","xls","xlsm"])
        _temp03.Level01 = JSON.parse(JSON.stringify(category_temp04))
		result[tempStr] = {
			"商品库存统计": _temp01,
			"营业占比分析上": _temp02,
			"营业占比分析下": _temp03,
		}
	}
	return result
}

function _getAloneInfo(classArea="2",fieldArea="F"){
	const wst = ThisWorkbook.Worksheets("操作")
	var result = {
		path: "",
		year: "",
		fields: {},
		subPaths: []
	}
	const name = wst.Cells(classArea,"A").Value2
	const path = wst.Cells(classArea,"B").Value2
	const year = wst.Cells(classArea,"C").Value2
	const usedRow = wst.Cells(wst.UsedRange.Rows.Count+256,fieldArea).End(xlUp).Row
	if (!folderExists(path)){
		MsgBox(`${name} 选中的路径不存在，请注意检查配置！`)
		return null
	}
	if (usedRow <= 1){
		MsgBox(`${name} 没有配置字段，请注意检查配置！`)
		return null
	}
	const arr = wst.Range(`${fieldArea}1`).Resize(usedRow,2).Value2
	for (let i = 1; i < usedRow;i++ ){
		var tempStr = arr[i][0]
		if (tempStr === "" || !tempStr){ continue }
		result.fields[tempStr] = {'column': arr[i][1] , 'columnIndex': toCol(arr[i][1])}
	}
	result.path = path
	result.year = year
	return result
}

function sortByColumn3(arr){
    var dataRows = []
    for (var i = 0; i < arr.length; i++){
        if (arr[i][0] !== "汇总" && arr[i][0] !== "合计"){
            dataRows.push({ origIndex: i, val: toVal(arr[i][2]) })
        }
    }
    dataRows.sort(function(a, b){
        var diff = b.val - a.val
        if (diff !== 0) return diff
        return a.origIndex - b.origIndex
    })
    var ranks = {}
    for (var j = 0; j < dataRows.length; j++){
        ranks[dataRows[j].origIndex] = j + 1
    }
    result = []
    for (var i = 0; i < arr.length; i++){
        if (arr[i][0] !== "汇总" && arr[i][0] !== "合计"){
            result.push([ranks[i],arr[i][1],arr[i][2]])
        }else{
            result.push([arr[i][0],arr[i][1],arr[i][2]])
        }
    }
    return result
}

function toCol(arg){
	try{
		return ThisWorkbook.Worksheets.Item(1).Cells(1,arg).Column
	}catch(error){
		return null
	}
}

function toVal(arg){
    try{
        return Number(arg)
    }catch(error){
        return 0
    }
}
/**
 * B2-1篮球/足球/排球 or A1-本册 拆分为 ["B2","1","篮球/足球/排球"] or ["A1","本册"]
 * @param {*} category 分类
 * @returns 分类后的数组
 */
function safeGet(obj, key1, key2, key3, key4){
    try{
        var temp = obj[key1]
        if (!temp) return 0
        temp = temp[key2]
        if (!temp) return 0
        temp = temp[key3]
        if (!temp) return 0
        if (key4 !== undefined){
            temp = temp[key4]
        }
        if (temp === undefined || temp === null) return 0
        return temp
    }catch(error){
        return 0
    }
}

function splitCategory(category){
    const idx = category.indexOf("-")
    if (idx <= 0){
        return ["","",category]
    }
    const first = category.slice(0, idx)
    const rest = category.slice(idx + 1)
    const match = rest.match(/^(\d+)(.+)$/)
    if (match){
        return [first, toVal(match[1]), match[2]]
    }
    return [first, "", rest]
}