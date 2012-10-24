#
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
#

require 'tables'

describe Tables::ExcelTableReader do

  it "should extract a table" do
    filename=File.dirname(__FILE__)+"/../test/rtm2.xlsx"
    @xtr = Tables::ExcelTableReader.new(filename)
    t=@xtr.extract_table("Requirements",{:rtf_columns=>[2], :progress=>true})

    t[7][2].should=="The message content should have the following fields.
 -  Message Id (Mandatory)
 -  Message Priority (Mandatory)
 -  Message Delivery Method (e.g. Portal, IHD ) (Mandatory)
 -  Message Type (e.g. Alert , General) (Mandatory)
 -  Message Text (Mandatory)
 -  Link(s) (embedded in text) (Optional)
 -  Message Start Timestamp (Mandatory)
 -  Message End timestamp (Mandatory)
 -  Recipient ID list (e.g. NMI) ( Mandatory)
 -  Group ID (Optional)
 -  Message Action (Mandatory)
 -  Sender (i.e. who is instigating the message)"

    t[5][2].should==""
    t[5][0].should=="CACP.OA.04"

    @xtr.exit
  end

end

