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

    office_use_only_row = rows.next
    resource_title_row = rows.next

    @resource_uri = get_or_create_resource(office_use_only_row, resource_title_row)

    if @resource_uri.nil?
      raise "No resource defined"
    end

    @headers = row_values(rows.next)

    begin
      while(row = rows.next)
        values = row_values(row)

        next if values.compact.empty?

        values_map = Hash[@headers.zip(values)]

        if values_map["Consignment"]
          @class_uri = get_or_create_class(values_map)
        end

        if values_map["Series Title"]
          @series_uri = get_or_create_series(values_map)
        end

        if values_map["Item Description"]
          add_file(values_map)
        end
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

  def get_or_create_resource(office_use_row, title_row)
    office_use_values = row_values(office_use_row)
    title_values = row_values(title_row)
    identifier_json = JSON(office_use_values[2,3] + [nil, nil])

    if (resource = Resource[:identifier => identifier_json])
      resource.uri
    else
      uri = "/repositories/12345/resources/import_#{SecureRandom.hex}"
      title = title_values[3]

      extent = {
        :portion => 'whole',
        :extent_type => 'metres',
        :container_summary => office_use_values[5],
        :number => office_use_values[4],
      }

      date = format_date(office_use_values[6])

      @records << JSONModel::JSONModel(:resource).from_hash({
                    :uri => uri,
                    :id_0 => office_use_values[2],
                    :id_1 => office_use_values[3],
                    :title => title,
                    :level => 'collection',
                    :extents => [extent],
                    :dates => [date].compact,
                    :language => 'eng',
                  })

      uri
    end
  end


  def get_or_create_class(row)
    return @class_uri if row['Consignment'].nil?

    class_hash = format_record(row).merge({
      :title => "Consignment #{row['Consignment']}",
      :level => 'class'
    })

    # consignments don't have these things
    [:dates, :instances, :component_id].map{|field| class_hash.delete(field)}

    @records << JSONModel::JSONModel(:archival_object).from_hash(class_hash)

    class_hash[:uri]
  end


  def get_or_create_series(row)
    return @series_uri if row['Series Title'].nil?

    series_hash = format_record(row).merge({
      :title => row['Series Title'],
      :level => 'series'
    })

    # if file defined in same row, do not add dates, instances
    # or component id to this series record
    if row["Item Description"]
      [:dates, :instances, :component_id].map{|field| series_hash.delete(field)}
    end

    series_hash[:parent] = { :ref => @class_uri } if @class_uri

    @records << JSONModel::JSONModel(:archival_object).from_hash(series_hash)

    series_hash[:uri]
  end


  def add_file(row)
    file_hash = format_record(row).merge({
      :title => row['Item Description'],
      :level => 'file'
    })

    file_hash[:parent] = { :ref => @class_uri } if @class_uri
    file_hash[:parent] = { :ref => @series_uri } if @series_uri

    @records << JSONModel::JSONModel(:archival_object).from_hash(file_hash)
  end


  def format_box(box_no, box_type)
    return if box_no.nil?

    {
      :instance_type => 'accession',
      :container => {
        :type_1 => box_type || 'Box',
        :indicator_1 => box_no
      }
    }
  end


  def format_date(date_string)
    return if date_string.nil?

    {
      :date_type => date_string =~ /-/ ? 'inclusive' : 'single',
      :label => 'creation',
      :expression => date_string || "No date provided"
    }
  end


  def row_values(row)
    (0...row.size).map {|i| (row[i] && row[i].value) ? row[i].value.to_s.strip : nil}
  end


  def format_record(row)
    record_hash = {
      :uri => "/repositories/12345/archival_objects/import_#{SecureRandom.hex}",
      :component_id => row['File no/ control no'],
      :instances => [format_box(row['Box No'], row['Box Type'])].compact,
      :dates => [format_date(row['Date Range'])].compact,
      :resource => {
        :ref => @resource_uri
      },
    }

    if row['Comments']
      record_hash['notes'] = [{
                                :jsonmodel_type => 'note_multipart',
                                :type => 'scopecontent',
                                :subnotes =>[{
                                               :jsonmodel_type => 'note_text',
                                               :content => row['Comments']
                                             }]
                              }]
    end

    record_hash
  end

end
