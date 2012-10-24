# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 

module Tables

  class TableReader

    attr_reader :tables

    def initialize
      @tables=[]
    end

    def get_tables
      result=[]
      if block_given? then
        @tables.each {|t| result<<t if yield(t)}
      else
        result=@tables.clone
      end
      result
    end

    def merge_tables
      sortmerge={}
      @tables.each do |t|
        key=t.header
        if sortmerge.has_key?(key) then
          sortmerge[key].merge!(t)
        else
          sortmerge[key]=t
        end
      end
      @tables=[]
      sortmerge.values.each {|v| @tables << v}
    end

    def extract_all_tables(idx_array=[*0..(table_count-1)]) # table_count is defined by the concrete class
      idx_array.each do |idx|
        table = extract_table(idx)
        unless table.nil? then
          status = ">> Extracted table '#{table.name}'"
          sig=">> signature - #{table.signature}"
          if block_given? then
            match=yield(table)
            if match then
              @tables << table
            else
              status += " - (filtered out)"
            end
          else
            @tables << table
          end
          puts status
          puts sig
        end
      end
    end

  end

end