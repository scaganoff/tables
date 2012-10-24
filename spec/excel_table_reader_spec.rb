# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 

require 'tables'

describe Tables::ExcelTableReader do

  it "should read an excel document" do
    filename=File.dirname(__FILE__)+"/../test/rtm.xlsx"
    @xtr = Tables::ExcelTableReader.new(filename)
    @xtr.workbook.should_not==nil
    @xtr.exit
  end

  it "should extract a table" do
    filename=File.dirname(__FILE__)+"/../test/rtm.xlsx"
    @xtr = Tables::ExcelTableReader.new(filename)
    t=@xtr.extract_table(0)
#    t.each_row {|r| puts "#{r[0]}, #{r[1]}, #{r[2]}"}
    t.rows.should==457
    t[1][0].should=~/Unique ID/
    t[2][0].should=~/CAMM\.AC\.101/
    t[1].last.should=~/PolicyNet Doc Ref/
  end

  it "should extract multiple tables" do
    filename=File.dirname(__FILE__)+"/../test/rtm.xlsx"
    @xtr = Tables::ExcelTableReader.new(filename)
    @xtr.extract_all_tables([0,1,2])
    @xtr.tables.count.should==3
  end

  it "should clean up a table" do
    filename=File.dirname(__FILE__)+"/../test/rtm.xlsx"
    @xtr = Tables::ExcelTableReader.new(filename)
    @xtr.extract_all_tables([2])
    @xtr.tables.count.should==1
    @xtr.clean
    t=@xtr.tables[0]
    t.rows.should==55

    col1=[]
    t.each_row {|row| col1<< row[0].strip }
    col1_expected=["Unique ID", "CAMM.GE.101", "CAMM.GE.102", "CAMM.GE.103", "CAMM.GE.104", "CAMM.GE.105", "CAMM.GE.106", "CAMM.GE.107", "CAMM.GE.108", "CAMM.GE.109", "CAMM.GE.110", "CAMM.GE.111", "CAMM.GE.112", "CAMM.GE.113", "CAMM.GE.114", "CAMM.GE.115", "CAMM.SM.101", "CAMM.SM.102", "CAMM.SM.103", "CAMM.BP.101", "CAMM.BP.102", "CAMM.BP.103", "CAMM.BP.104", "CAMM.BP.105", "CAMM.BP.106", "CAMM.BP.107", "CAMM.BP.108", "CAMM.GP.101", "CAMM.SI.101", "CAMM.SI.102", "CAMM.HW.101", "CAMM.HW.102", "CAMM.HW.103", "CAMM.HW.104", "CAMM.HA.101", "CAMM.HA.102", "CAMM.HA.103", "CAMM.HA.104", "CAMM.HA.105", "CAMM.HA.106", "CAMM.HA.107", "CAMM.HA.108", "CAMM.HA.109", "CAMM.HA.110", "CAMM.HA.111", "CAMM.HA.112", "CAMM.AI.101", "CAMM.AI.102", "CAMM.AI.103", "CAMM.WG.101", "CAMM.WG.102", "CAMM.WG.103", "CAMM.WG.104", "CAMM.WG.105", "CAMM.MU.101"]
    col1.should.eql? col1_expected
  end

  it "should write a table" do
    t=Tables::Table.new([
        [:a,:b,:c,:d],
        [1,1,1,1],
        [2,2,2,2],
        [3,3,3,3]
      ])
    outfile=File.dirname(__FILE__)+"/../test/writetest.xlsx"
    File.delete(outfile) if File.exist?(outfile)
    @xtr=Tables::ExcelTableReader.new()
    @xtr.create_file(outfile)
    @xtr.write_table(t)
    @xtr.save
    @xtr.exit
  end

  it "should write a table to a named worksheet" do
    t=Tables::Table.new([
        [:a,:b,:c,:d],
        [1,1,1,1],
        [2,2,2,2],
        [3,3,3,3]
      ])
    outfile=File.dirname(__FILE__)+"/../test/writetest1.xlsx"
    File.delete(outfile) if File.exist?(outfile)
    @xtr=Tables::ExcelTableReader.new()
    @xtr.create_file(outfile)
    @xtr.write_table(t, "Sheet2")
    @xtr.save
    @xtr.exit
  end

  it "should write a column" do
    t=Tables::Table.new([
    [:a,:b,:c,:d],
    [1,1,1,1],
    [2,2,2,2],
    [3,3,3,3]
    ])
    outfile=File.dirname(__FILE__)+"/../test/writetest2.xlsx"
    File.delete(outfile) if File.exist?(outfile)
    @xtr=Tables::ExcelTableReader.new()
    @xtr.create_file(outfile)
    @xtr.write_column(t,:c)
    @xtr.save
    @xtr.exit
  end

  it "should write a column to a named worksheet" do
    t=Tables::Table.new([
    [:a,:b,:c,:d],
    [1,1,1,1],
    [2,2,2,2],
    [3,3,3,3]
    ])
    outfile=File.dirname(__FILE__)+"/../test/writetest2a.xlsx"
    File.delete(outfile) if File.exist?(outfile)
    @xtr=Tables::ExcelTableReader.new()
    @xtr.create_file(outfile)
    @xtr.write_column(t,:c, "Sheet3")
    @xtr.save
    @xtr.exit
  end

  it "should append worksheets automatically" do
    t=Tables::Table.new([
    [:a,:b,:c,:d],
    [1,1,1,1],
    [2,2,2,2],
    [3,3,3,3]
    ])
    outfile=File.dirname(__FILE__)+"/../test/writetest3.xlsx"
    File.delete(outfile) if File.exist?(outfile)
    @xtr=Tables::ExcelTableReader.new()
    @xtr.create_file(outfile)
    10.times {|idx| @xtr.write_table(t,idx)}
    @xtr.save
    @xtr.exit
  end

end

