#
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
#
require "rspec"
require 'tables'

describe Tables::WordTableWriter do

  before(:each) do
    @wtw=Tables::WordTableWriter.new
    @table=Tables::Table.new([
        ["Header 1", "Header 2", "Header 3"],
        ["Cell(2,1)","Cell(2,2)","Cell(2,3)"],
        ["Cell(3,1)","Cell(3,2)","Cell(3,3)"],
        ["Cell(4,1)","Cell(4,2)","Cell(4,3)"],
        ["Cell(5,1)","Cell(5,2)","Cell(5,3)"],
        ["Cell(6,1)","Cell(6,2)","Cell(6,3)"]
    ])
    @outfile=File.dirname(__FILE__)+"/../test/table_writer_test.docx"
    begin; File.delete(@outfile); rescue; end
  end

  it "should write a table to word" do
    File.exists?(@outfile).should==false
    @wtw.append_table(@table, "This is my very nice table.")
    @wtw.save_as(@outfile)
    @wtw.exit
    File.exists?(@outfile).should==true

    @wtr=Tables::WordTableReader.new(@outfile)
    t_result=@wtr.extract_table(0)
    t_result.should.eql? @table
  end

end