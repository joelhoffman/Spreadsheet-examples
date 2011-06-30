require 'rubygems'
require 'spreadsheet'

['oo-unsecure.xls', 'oo-secure.xls', 'ms-unsecure.xls', 'ms-secure.xls'].each do |f|
  begin
    Spreadsheet.open(f)
    puts "Successfully opened %s" % f
  rescue RuntimeError => e
    puts "Failure opening %s: %s" % [f, e.message]
  end
end

