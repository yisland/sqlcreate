require 'yaml'
require 'spreadsheet'

SETTING = YAML.load_file("settings.yml")
TMPFILEPATH = SETTING["default"]["templateFilePath"]
OUTPUTPATH = SETTING["default"]["outputFilePath"]
EXCLUSION = SETTING["default"]["exclusion"]

def cellFormat(cell)
    cell = cell.value  if cell.instance_of?(Spreadsheet::Formula)
    if cell.instance_of?(String) then
        flg = false
        EXCLUSION.each do |ex|
              flg = true if cell.casecmp(ex) == 0
        end
        cell = "'" + cell + "'" if flg == false
    end
    return cell
end

def insert(name, columnArray, cellArray)
    sql = "INSERT INTO " + name + "(" 
    columnArray.each do |column|
        sql << column + ","
    end
    sql.chop!
    sql << ") VALUES("
    cellArray.each do |cell|
        if !cell.instance_of?(String) then
            sql.concat(cell.to_s + ",")
        else
            sql.concat(cell + ",")
        end
    end
    sql.chop!
    sql << ");"
    return sql
end

def makeFile(sqlArray, name)
    fileName = OUTPUTPATH + "/" + name + ".sql"
    FileUtils.mkdir_p(File::dirname(fileName)) unless FileTest.exists?(fileName)
    File.open(fileName, "w") do |f|
        sqlArray.each do |item|
            f.write item + "\n"
            puts item
        end
    end
end

Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet.open(TMPFILEPATH, 'rb')

columnArray = []
cellArray = []
sqlArray = []

begin
    book.worksheets.each do |ws|
        sqlArray.clear
        columnArray.clear
        ws.each_with_index do |row, row_idx|
            cellArray.clear
            row.each do |cell|
                if row_idx == 0 then
                    columnArray.push cell
                    next
                end
                cellArray.push cellFormat(cell)
            end
            if row_idx > 0 then
                if cellArray.length != 0 then
                    sqlArray.push insert(ws.name, columnArray, cellArray)
                end
            end
        end
        makeFile(sqlArray, ws.name)
    end
rescue => ex
    puts "Error " + ex.message
    puts ex.backtrace
    exit
end