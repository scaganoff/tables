# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 

require 'tables'

describe Tables::WordTableReader do

  it "should read a word document" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    @wtr.doc.should_not==nil
    @wtr.exit
  end

  it "should count the tables" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    @wtr.doc.should_not==nil
    @wtr.doc.tables.count.should==55
    @wtr.exit
  end

  it "should extract a table" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    t=@wtr.extract_table(21)
    t.rows.should==13
    t[0][0].should=~/Unique ID/
    t[1][0].should=~/CAMM\.SC\.101/
  end

  it "should extract all tables" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    @wtr.extract_all_tables((0..53).to_a) # TODO: Problem extracting last table should be 0..55
    @wtr.tables.count.should==54
  end

  it "should filter tables on extract" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    @wtr.extract_all_tables((0..53).to_a) {|table| table[0][0]=~/Unique ID/} # TODO: Problem extracting last table should be 0..55
    @wtr.tables.count.should==35
  end

  it "should merge similar tables" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    @wtr.extract_all_tables((0..53).to_a) # TODO: Problem extracting last table should be 0..55
    @wtr.merge_tables
    @wtr.tables.count.should==19
  end

  it "should return tables matching a filter" do
    filename=File.dirname(__FILE__)+"/../test/rtm.docx"
    @wtr = Tables::WordTableReader.new(filename)
    @wtr.extract_all_tables((0..53).to_a) # TODO: Problem extracting last table should be 0..55
    @wtr.merge_tables
    tables=@wtr.get_tables {|table| table[0][0]=~/Unique ID/}
    tables.count.should==1
    column1=[]
    tables[0].each_row {|row| column1<< row[0] }
    column1.count.should==382
  end

end

