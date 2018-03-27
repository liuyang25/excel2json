package.cpath=package.cpath..";../?.dll;./?.dll;../tool/?.dll;"
require ("lfs")
require ("lc")

-- print(lc.help())

FIX_RANGE = 100 --range一次性加载会报错，只能分段加载

function excel2Lua(excelFile)
	require ("luacom")	--确保加载的是同一级目录的那个

	--打开表格文件
	local iExcel = luacom.CreateObject("Excel.Application")
	-- local iExcel = luacom.CreateObject("Ket.Application")
	local curDir = lfs.currentdir()
	print("curDir:"..curDir)
	local fileName = excelFile or (curDir .. '/' ..F(lc.u2a("员工贷常见问题.xlsx")))
	print("open.."..excelFile)
	iExcel.WorkBooks:Open(fileName, nil, 0)

	readExcel(iExcel)

	iExcel.Application:Quit()
end


-- ****************************************************************** --

local FIX_RANGE = 100
local dataOutput = {}

function Format(s) 
	s = string.gsub(s, '"', '\\"')	
	s = string.gsub(s, '\n', '\\n')	
	return s
end
function readExcel(iExcel)
	local iSheet = iExcel.Worksheets(1)
	io.write("\n")
	local usedrange = iSheet.UsedRange
	local linecount = usedrange.Rows.Count

	local range = usedrange.Rows("1:"..math.min(FIX_RANGE, linecount)).Value2
	for line = 2, linecount do
		if line % FIX_RANGE == 1 then
			range = usedrange.Rows(line..":"..math.min(line+FIX_RANGE-1, linecount)).Value2
		end

		local rowData = range[((line - 1) % FIX_RANGE) + 1]

		local q = rowData[1]
		local a = rowData[2]
		if not q then break end
		
		q = Format(q)
		a = Format(a)
		local data = '  { "q": "' .. tostring(q) ..'", "a": "'.. tostring(a)..'" }'

		-- next line null, this ends none
		if not iSheet.Cells(line + 1, 1).Value2 then
			data = data .. '\n'
		else
			data = data .. ',\n'
		end
		dataOutput[#dataOutput + 1] = data
	end

	local outputStr = table.concat(dataOutput, '')
	-- A2U(outputStr)
	-- U2A(outputStr)
	outputStr = F(lc.a2u(outputStr))

	--output
	local outputFile = './FAQ.json'

	os.execute("if not exist "..outputFile.." mkdir "..outputFile)

	local file = io.open(outputFile, "w")
	if file == nil then print("open failed: "..outputFile) return end

	local fileData = '[\n' .. outputStr ..']\n'
	file:write(fileData)
	file:close()
	print("hehe")
end

function F(str)
	if str == nil then return nil end
	str = string.sub(str, 1, string.len(str) - 1)
	return str
end

--执行脚本
excel2Lua(arg[1])

--os.execute("pause")


