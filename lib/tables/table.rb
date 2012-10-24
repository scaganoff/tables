# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 

module Tables

  class Table

    attr_reader :table, :colindex, :idcolumn, :rowindex
    attr_accessor :name

    def initialize(table=nil,name=nil)
      @table=[]
      table.each {|row| @table<<row } unless table.nil?
      build_column_index unless table.nil?
      self.name=name unless name.nil?
    end

    def header
      @table[0]
    end

    #def columns
    #  self.header.count
    #end

    def idcolumn=(column_name)
      raise "Unknown column '#{column_name}'" unless @colindex.has_key?(column_name)
      @idcolumn=column_name
      build_row_index
    end

    def get_value(column_name,row_num)
      col_num=@colindex[column_name]
      raise "Unknown column name '#{column_name}'" if col_num.nil?
      @table[row_num][col_num]
    end

    def get_row(arg)
      result=nil
      if arg.is_a? Integer then
        result=get_row_by_num(arg)
      else
        result=get_row_by_num(@rowindex[arg])
      end
    end

    def add_row(row)
      add_row_array(row) if row.is_a?(Array)
      add_row_hash(row) if (row.is_a?(Hash) and self.row_count>0)
      add_first_row_hash(row) if (row.is_a?(Hash) and self.row_count==0)
    end

    def [](idx)
      @table[idx]
    end

    def ==(other)
      self.table==other.table
    end

    # obsolescent
    def rows
      puts "WARNING: use 'row_count' instead of 'rows'"
      self.row_count
    end

    def column_count
      self.header.count
    end

    def row_count
      @table.count
    end

    def similar?(t2)
      self.header==t2.header
    end

    def merge!(t2)
      raise "tables are not similar" unless similar?(t2)
      (1..t2.row_count-1).each do |idx|
        begin
          self<<t2[idx]
        rescue Exception=>e
          puts "ERROR: Error adding row #{idx} from '#{t2.name}' to '#{self.name}'"
          puts "ERROR: #{e.message}"
        end
      end
    end

    # << is a synonym of add_row
    def <<(row)
      self.add_row(row)
    end

    # TODO: Make each_row behave like get_each_row
    def each_row
      @table.each {|row| yield(row)}
    end

    def each_row_with_index
      @table.each_with_index {|row,idx| yield(row,idx) }
    end

    def get_each_row(skip_header=true)
      @table.each_with_index {|row,idx| yield(self.get_row(idx)) unless (idx==0 and skip_header)}
    end

    # TODO: How can we do this within the get_row closure?
    def set_row(row)
      idx=row[:rownum]
      row.each_pair do |colname,value|
        jdx=@colindex[colname]
        @table[idx][jdx]=value unless jdx.nil?
      end
    end

    def remove_blank_rows!(startcol=0)
      remove_matched_rows! { |row| row[startcol..-1].join.strip=="" }
    end

    def remove_repeat_headers!
      header=self.header
      remove_matched_rows! { |row| row.eql?(header) and not(row.equal?(header)) }
    end

    def remove_matched_rows!
      blanks=[]
      @table.each do |row|
        blanks << row if yield(row)
      end
      blanks.each {|br| @table.delete_if {|el| el.equal?(br) }} # delete if objects are the same
      build_column_index
      return self
    end

    def demerge!(colnum=0)
      new_table=[@table[0]]
      (1..@table.count-1).each do |idx|
        demerge_it(colnum,idx, new_table)
      end
      @table=new_table
      return self
    end

    # Fill blank cells in a column with the last non-blnak value. This is aimd at merged excel cells
    # where all cells have the same value, but merging sets all cells blank except for the top-left
    def fill_column_down(column_name, ignore_header=true)
      colnum=@colindex[column_name]
      last_value=""
      start = ignore_header ? 1:0
      (start..self.row_count-1).each do |idx|
        value=@table[idx][colnum]
        if value == ""
          @table[idx][colnum] = last_value
        else
          last_value = value
        end
      end
    end

    def rename_column(old_name, new_name)
      colnum=@colindex[old_name]
      raise "Unknown column '#{name}'" if colnum.nil?
      @colindex[new_name]=colnum
      @colindex.delete(old_name)
      @table[0][colnum]=new_name
    end

    def delete_column(name)
      colnum=@colindex[name]
      raise "Unknown column '#{name}'" if colnum.nil?
      self.each_row {|row| row.delete_at(colnum)}
      @colindex.delete(name)
      @colindex.each_pair {|k,v| @colindex[k]=v-1 if v>colnum }
    end

    def get_column(name)
      colnum=@colindex[name]
      raise "Unknown column '#{name}'" if colnum.nil?
      result=[]
      self.each_row {|row| result<<row[colnum]}
      result
    end

    def column_copy(other_table)
      other_table.get_each_row do |other_row|
        self.add_row(other_row)
      end
    end

    def signature
      self.header.join(',')
    end

  private

    def add_row_array(row)
      raise "Argument must be an array" unless row.is_a? Array
      unless self.header.nil? then
        n=self.header.count
        raise "Row '#{row[0]}' must have #{n} values...found only #{row.count}" unless row.count==n
      end
      @table<<row
      build_column_index if @table.count==1
    end

    def add_row_hash(row)
      raise "Argument must be a hash table" unless row.is_a? Hash
      n=self.header.count unless self.header.nil?
      new_row=Array.new(n)
      row.each_pair do |k,v|
        idx=@colindex[k]
        new_row[idx]=v unless idx.nil?
      end
      @table<<new_row
    end

    def add_first_row_hash(row)
      raise "Argument must be a hash table" unless row.is_a? Hash
      new_header=[]
      new_row=[]
      row.each_pair do |k,v|
        new_header<<k
        new_row<<v
      end
      add_row_array(new_header)
      add_row_array(new_row)
    end

    def build_row_index
      @rowindex={}
      self.get_each_row do |row|
        id=row[@idcolumn]
        @rowindex[id]=row[:rownum]
      end
    end

    def build_column_index
      @colindex={}
      self.header.each_with_index {|value,idx| @colindex[value]=idx }
    end

    def get_row_by_num(row_num)
      row=@table[row_num]
      result={}
      row.each_with_index {|v,idx| result[self.header[idx]]=v }
      result[:rownum]=row_num
      result
    end

    def demerge_it(colnum,idx,new_table)
      next_row=@table[idx]
      if next_row[colnum].strip=="" then
        row=new_table.pop
        demerged_row=demerge_two_rows(row, next_row)
        new_table.push(demerged_row)
      else
        new_table.push(next_row)
      end
    end

    def demerge_two_rows(r1, r2)
      raise "Column number mismatch" if r1.count != r2.count
      new_row=[]
      (0..r1.count-1).each do |idx|
        new_cell=r1[idx]+"\n"+r2[idx]
        new_row << new_cell.strip
      end
      new_row
    end

  end

end