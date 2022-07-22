# Скрипт для вычленения имён и телефонов абонентов из служебных файлов

require 'roo'
require 'write_xlsx'

file_names = []
person_list = []

# Собираем имена файлов в папке в массив
Dir.chdir('input_files') do
	file_names = Dir.glob('*.xlsx').each { |file_name| file_name }
end

file_names.each do |file_name|
	xlsx = Roo::Spreadsheet.open("./input_files/#{file_name}")

	# Получим массив строк столбца имён, начиная с 6ой строки
	name_list = xlsx.sheet(0).column(17)[5..].each do |row|
		row
	end

	# Получим массив строк столбца номеров, начиная с 6ой строки
	phone_list = xlsx.sheet(0).column(18)[5..].each do |row|
		row
	end

	# Получив данные в переменные name_list, phone_list, закрываем документ
	xlsx.close

	# Делаем двумерный массив по парам [[имя, номер], ...]
	# для удобства записи построчно через итерацию при помощи гема write_xlsx
	name_list.each_with_index do |name, index|
		person = [name, phone_list[index] ]
		person_list << person
	end
end
# Делаем список уникальным
person_list.uniq!

# Создаём экземпляр книги и листа
workbook = WriteXLSX.new("file_name.xlsx")
worksheet = workbook.add_worksheet

# Записываем построчно на лист
person_list.each.with_index(1) do |row, index|
	p row # ['имя', 'номер']
	worksheet.write_row(index, 0, row)
end

workbook.close
