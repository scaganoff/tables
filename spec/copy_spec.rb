# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 

require "rspec"
require "tables"

describe Tables::Table do

  it "should do a columnwise copy of one table to another" do
    t1=Tables::Table.new([
      [:a,:b,:d,:f],
      [1,1,1,1],
      [2,2,2,2],
      [3,3,3,3]
      ])
    t2=Tables::Table.new([
      [:a,:b,:c,:d,:e,:f]
      ])
    t_expect=Tables::Table.new([
      [:a,:b,:c,:d,:e,:f],
      [1,1,nil,1,nil,1],
      [2,2,nil,2,nil,2],
      [3,3,nil,3,nil,3]
      ])

    t2.column_copy(t1)
    (t2==t_expect).should==true
  end

end