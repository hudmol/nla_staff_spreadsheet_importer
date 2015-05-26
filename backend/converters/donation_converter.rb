require 'date'
require 'rubyXL'


class DonationConverter < Converter

  def self.instance_for(type, input_file)
    if type == "donation"
      self.new(input_file)
    else
      nil
    end
  end


  def self.import_types(show_hidden = false)
    [
      {
        :name => "donation",
        :description => "Donor Box List spreadsheet"
      }
    ]
  end


  def self.profile
    "Convert an Donor spreadsheet to ArchivesSpace JSONModel records"
  end


  def initialize(input_file)
    super
    @batch = ASpaceImport::RecordBatch.new
    @input_file = input_file
    @records = []
  end


  def run
    workbook = RubyXL::Parser.parse(@input_file)

    sheet = workbook[0]

    rows = sheet.enum_for(:each)

    rows.next # skip "Office Use Only" row

    resource_identifier_row = rows.next
    resource_title_row = rows.next

    @resource_uri = get_or_create_resource(resource_identifier_row, resource_title_row)

    if @resource_uri.nil?
      raise "No resource defined"
    end

    @headers = row_values(rows.next)

    begin
      while(row = rows.next)
        values = row_values(row)

        next if values.compact.empty?

        @series_uri = get_or_create_series(values[2])

        if @series_uri.nil?
          raise "No series defined for item: #{values}"
        end

        add_item(Hash[@headers.zip(values)])
      end
    rescue StopIteration
    end

    # assign all records to the batch importer in reverse
    # order to retain position from spreadsheet
    @records.reverse.each{|record| @batch << record}
  end


  def get_output_path
    output_path = @batch.get_output_path

    p "=================="
    p output_path
    p File.read(output_path)
    p "=================="

    output_path
  end


  private

  def get_or_create_resource(identifier_row, title_row)
    identifier_values = row_values(identifier_row)
    identifier_json = JSON(identifier_values[1,4])

    if (resource = Resource[:identifier => identifier_json])
      resource.uri
    else
      uri = "/repositories/12345/resources/import_#{SecureRandom.hex}"
      title = row_values(title_row)[1]

      @records << JSONModel::JSONModel(:resource).from_hash({
                    :uri => uri,
                    :id_0 => identifier_values[1],
                    :id_1 => identifier_values[2],
                    :id_2 => identifier_values[3],
                    :id_3 => identifier_values[4],
                    :title => title,
                    :level => 'collection',
                    :extents => [{
                      :portion => 'whole',
                      :number => '0',
                      :extent_type => 'linear_feet'
                    }],
                    :dates => [{
                      :date_type => 'single',
                      :label => 'other',
                      :expression => "Imported from donor spreadsheet on #{Date.today}"
                    }]
                  })

      uri
    end
  end


  def add_item(hash)
    @records << JSONModel::JSONModel(:archival_object).from_hash({
                  :uri => "/repositories/12345/archival_objects/import_#{SecureRandom.hex}",
                  :level => 'item',
                  :title => hash['Item Description'],
                  :instances => [{
                    :instance_type => 'mixed_materials',
                    :container => {
                      :type_1 => 'box',
                      :indicator_1 => hash['Box No'],
                      :type_2 => 'folder',
                      :indicator_2 => hash['File no/ control no']
                    }
                  }],
                  :dates => [{
                    :date_type => hash['Date Range'] =~ /-/ ? 'inclusive' : 'single',
                    :label => 'existence',
                    :expression => hash['Date Range'] || "No date provided"
                  }],
                  :collection_management => {
                    :cataloged_note => hash['Comments']
                  },
                  :resource => {
                    :ref => @resource_uri
                  },
                  :parent => {
                    :ref => @series_uri
                  }
                })
  end


  def row_values(row)
    (0...row.size).map {|i| (row[i] && row[i].value) ? row[i].value.to_s.strip : nil}
  end


  def parse_identifier(collection_id)
    JSON((collection_id.split(/\//) + [nil, nil, nil, nil]).take(4))
  end



  def get_or_create_series(title)
    return @series_uri if title.nil?

    uri = "/repositories/12345/archival_objects/import_#{SecureRandom.hex}"

    @records << JSONModel::JSONModel(:archival_object).from_hash({
                 :title => title,
                 :level => 'series',
                 :uri => uri,
                 :resource => {
                   :ref => @resource_uri
                 }
               })

    uri
  end
end