require 'yaml'
require 'spreadsheet'

SETTING = YAML.load_file("settings.yml")
TMPFILEPATH = SETTING["default"]["templateFilePath"]
OUTPUTPATH = SETTING["default"]["outputFilePath"]
EXCLUSION = SETTING["default"]["exclusion"]
EXCLUSIONSHEET = SETTING["default"]["exclusionSheet"]

def cellFormat(cell)
    cell = cell.value  if cell.instance_of?(Spreadsheet::Formula)
    if cell.instance_of?(DateTime) then
        cell = cell.strftime("%Y/%m/%d %X")
    end
    if cell.instance_of?(String) then
        cell = "'" + cell + "'" if !EXCLUSION.include?(cell)
    end
    return cell
end

def insert(name, hash)
    sql = "INSERT INTO " + name + "(" 
    hash.each_key do |key|
        sql << key + ","
    end
    sql.chop!
    sql << ") VALUES("
    hash.each_value do |value|
        if !value.instance_of?(String) then
            sql.concat(value.to_s + ",")
        else
            sql.concat(value + ",")
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

columnArray = []
cellArray = []
sqlArray = []

begin
    book = Spreadsheet.open(TMPFILEPATH, 'rb')

    book.worksheets.each do |ws|
        sqlArray.clear
        columnArray.clear
        next if EXCLUSIONSHEET.include?(ws.name)
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
                next if columnArray.length != cellArray.length
                hash = {}
                columnArray.each_with_index do |columnRow, idx|
                    cell = cellArray[idx]
                    hash.store(columnRow, cell) if cell != nil
                end
                if hash.length != 0 then
                    sqlArray.push insert(ws.name, hash)
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