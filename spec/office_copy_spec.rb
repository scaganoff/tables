# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 

require "rspec"
require 'tables'

describe Tables::WordTableReader do

  it "should copy word table to excel" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    ntables=@wtr.table_count
    @wtr.extract_all_tables((0..ntables-2).to_a) {|table| table[0][0]=~/Unique ID/}
    @wtr.merge_tables
    @wtr.tables.should have(1).item
    table=@wtr.tables[0]
    @wtr.exit

    outfile=File.dirname(__FILE__)+"/../test/copytest.xlsx"
    File.delete(outfile) if File.exist?(outfile)
    @xtr=Tables::ExcelTableReader.new()
    @xtr.create_file(outfile)
    @xtr.write_table(table)
    @xtr.save
    @xtr.exit
  end

  it "should copy word table to an excel template" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    ntables=@wtr.table_count
    @wtr.extract_all_tables((0..ntables-2).to_a) {|table| table[0][0]=~/Unique ID/}
    @wtr.merge_tables
    @wtr.tables.should have(1).item
    table=@wtr.tables[0]
    @wtr.exit

    template_file=File.dirname(__FILE__)+"/../test/rtm_template.xlsx"
    ttr=Tables::ExcelTableReader.new(template_file)
    template=ttr.extract_table(1)
    template.column_copy(table)
    ttr.exit

    outfile=File.dirname(__FILE__)+"/../test/copytest2.xlsx"
    File.delete(outfile) if File.exist?(outfile)
    @xtr=Tables::ExcelTableReader.new()
    @xtr.create_file(outfile)
    @xtr.write_table(template)
    @xtr.save
    @xtr.exit
  end
end