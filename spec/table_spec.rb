# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 

require 'tables'

describe Tables::Table do
  before(:each) do
    @t1=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      [2,2,9,2],
      [3,3,3,3]
    ])
    @t2=Tables::Table.new([
      [:a,:b,:c,:d],
      [4,4,4,4],
      [5,5,5,5]
    ])
    @t3=Tables::Table.new([
      [:x,:y,:z],
      [4,4,4],
      [5,5,5]
    ])
    @texpect=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      [2,2,9,2],
      [3,3,3,3],
      [4,4,4,4],
      [5,5,5,5]
    ])
  end

  it "should determine equality" do
    t=Tables::Table.new([
        [:a,:b,:c,:d],
        [1,1,1,1],
        [2,2,9,2],
        [3,3,3,3]
     ])
    (t==@t1).should==true
  end
  it "should determine similarity" do
    @t1.similar?(@t2).should==true
    @t1.similar?(@t3).should==false
  end

  it "should merge tables cleanly" do
    @t1.merge!(@t2)
    (@t1==@texpect).should==true
  end

  it "should accept a new row" do
    @t1 << [4,4,4,4]
    @t1 << [5,5,5,5]
    (@t1==@texpect).should==true
  end

  it "should report columns" do
    @t1.column_count.should==4
  end

  it "should report rows" do
    @t1.row_count.should==4
  end

  it "should access cell value by column name" do
    @t1.get_value(:c, 2).should==9
    lambda { t.get_value(:z,2) }.should raise_error
  end

  it "should provide row as a hash" do
    row=@t1.get_row(2)
    row.should.is_a? Hash
    row[:c].should==9
    row[:a].should==2
  end

  it "should get a named column" do
    c=@t1.get_column(:c)
    c.should.eql? [:c,1,9,3]
  end

  it "should delete a named column" do
    c=@t1.get_column(:c)
    c.should.eql? [:c,1,9,3]
    t=@t1.delete_column(:c)
    t_expect=Tables::Table.new([
      [:a,:b,:d],
      [1,1,1],
      [2,2,2],
      [3,3,3]
    ])
    lambda { t.get_column(:c) }.should raise_error
  end

  it "should delete the last column" do
    @t1.delete_column(@t1.header.last)
    texpect=Tables::Table.new([
      [:a,:b,:c],
      [1,1,1],
      [2,2,9],
      [3,3,3]
    ])
    (@t1==texpect).should==true
  end

  it "should remove blank rows" do
    t_blankrow=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      ["","","",""],
      [3,3,3,3]
    ])
    t_noblankrow=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      [3,3,3,3]
    ])
    t_blankrow.remove_blank_rows!
    t_blankrow.row_count.should==3
    (t_blankrow==t_noblankrow).should==true
  end

  it "should remove head rows" do
    t_blankrow=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      ["head","","",""],
      [3,3,3,3]
    ])
    t_noblankrow=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      [3,3,3,3]
    ])
    t_blankrow.remove_blank_rows!(1)
    (t_blankrow==t_noblankrow).should==true
    t_blankrow.row_count.should==3
  end

  it "should remove repeated headers" do
    t=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      [2,2,2,2],
      [3,3,3,3],
      [:a,:b,:c,:d],
      [4,4,4,4],
      [5,5,5,5]
    ])
    t_expect=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,1,1],
      [2,2,2,2],
      [3,3,3,3],
      [4,4,4,4],
      [5,5,5,5]
    ])
    t.remove_repeat_headers!
    (t==t_expect).should==true
    t.row_count.should==6
  end

  it "should demerge on default column 0" do
    t_merged=Tables::Table.new([
      ["a","foo","45"],
      ["","sub-foo","46"],
      ["","sub-foo-2",""],
      ["b","bar","47"],
      ["c","baz",""]
    ])
    t=t_merged.demerge!
    t_expect=Tables::Table.new([
      ["a","foo\nsub-foo\nsub-foo-2","45\n46"],
      ["b","bar","47"],
      ["c","baz",""]
      ])
    (t==t_expect).should==true
  end

  it "should demerge on any specified column" do
    t_merged=Tables::Table.new([
      ["a","foo","45"],
      ["","sub-foo","46"],
      ["","sub-foo-2",""],
      ["b","bar","47"],
      ["c","baz",""]
    ])
    t=t_merged.demerge!(2)
    t_expect=Tables::Table.new([
      ["a","foo","45"],
      ["","sub-foo\nsub-foo-2","46"],
      ["b\nc","bar\nbaz","47"]
      ])
    (t==t_expect).should==true
  end

  it "should allow build Tables::Table" do
    t=Tables::Table.new
    t.add_row([:a,:b,:c,:d])
    t.add_row([1,1,1,1])
    t.add_row([2,2,9,2])
    t.add_row([3,3,3,3])
    (t==@t1).should==true
    t.get_value(:c,2).should==9
  end

  it "should allow rename of column" do
    t=Tables::Table.new
    t.add_row([:a,:b,:c,:d])
    t.add_row([1,1,1,1])
    t.add_row([2,2,9,2])
    t.add_row([3,3,3,3])
    t.rename_column(:c, :z)
    t.get_value(:z,2).should==9
    lambda { t.get_value(:c,2) }.should raise_error
  end

  it "should update a row via get" do
    t_expect=Tables::Table.new([
      [:a,:b,:c,:d],
      [1,1,9,1],
      [2,2,9,2],
      [3,3,9,3]])
    @t1.get_each_row do |row|
      row[:c]=9
      @t1.set_row(row)
    end
    (@t1==t_expect).should==true
  end

end

