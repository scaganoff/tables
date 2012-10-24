# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 
require 'win32ole'

module Tables

  class WordTableReader < TableReader

    attr_reader :doc

    def initialize(filename, opts={})
      fso=WIN32OLE.new('Scripting.FileSystemObject')
      path=fso.GetAbsolutePathName(filename)
      @app=WIN32OLE.new("Word.Application")
      @doc=@app.Documents.Open(path, opts)
      super()
    end

    def table_count
      @doc.tables.count
    end

    def extract_table(idx)
      name="Table #{idx+1}"
      table=@doc.tables(idx+1)
      result=[]
      jdx=0
      table.rows.each do |row|
        jdx+=1
        begin
          result << WordTableReader.extract_row(row)
        rescue Exception=>e1
          puts "ERROR: Error extracting row #{jdx} from table '#{name}'"
          puts "ERROR: #{e1.message}"
        end
      end
      Table.new(result, name)

    rescue Exception=>e2
      puts "ERROR: Error extracting table '#{name}'"
      puts "ERROR: #{e2.message}"
    end

    def exit
      @app.quit
    end

  private

    def WordTableReader.extract_row(ole_row)
      row=[]
      ole_row.cells.each do |cell|
        row << WordTableReader.extract_text(cell.range)
      end
      row
    end

    def WordTableReader.extract_text(range)
      paragraphs=[]
      range_text=range.Text
      range.paragraphs.each do |p|
        text = clean_bytes(p.range.Text)
        list_text = clean_bytes(p.range.ListFormat.ListString)
        list_type = p.range.ListFormat.ListType
        p_text = text.empty? ? list_text : text   # return list_text if text is empty
        p_bullet = list_type==0 ? "" : "\t- "
        paragraphs << p_bullet+p_text.strip
      end
      paragraphs.join("\n")
    end

    def WordTableReader.clean_bytes(string)
      new_string=""
      string.each_char {|c| new_string<<c unless c=="\r" or c=="\a"}
      new_string
    end

  end

end