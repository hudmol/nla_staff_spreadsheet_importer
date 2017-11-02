class BasicResourceConverter < Converter

  def self.instance_for(type, input_file)
    if type == "basic_resource"
      self.new(input_file)
    else
      nil
    end
  end


  def self.import_types(show_hidden = false)
    [
      {
        :name => "basic_resource",
        :description => "Paper Collection Sheets CSV"
      }
    ]
  end


  def self.profile
    "Convert a Paper Collection Sheets CSV to ArchivesSpace Resource records"
  end


  def initialize(input_file)
    super
    @batch = ASpaceImport::RecordBatch.new
    @input_file = input_file
    @records = []

    @columns = %w(
                  title
                  resource_id
                  access_conditions
                  use_conditions
                  scopecontent_note
                  processing_note_1
                  processing_note_2
                  date_expression
                  extent_container_summary
                  extent_number
                  extent_type
                 )

  end


  def run
    rows = CSV.read(@input_file)

    begin
      while(row = rows.shift)
        values = row_values(row)

        next if values.compact.empty?

        values_map = Hash[@columns.zip(values)]

        # skip header rows
        next if values_map['title'].nil? ||
                values_map['title'].strip == '' ||
                values_map['title'] == 'resources_basicinformation_title' ||
                values_map['title'] == 'Title'

        create_resource(values_map)
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

  def create_resource(row)
    # turns out Emma wants the whole id in id_0
    # leaving this stuff here because, when, you know ...
    # id_a = row['resource_id'].split(/\s+/)
    id_a = [row['resource_id']]
    id_a = id_a + Array.new(4 - id_a.length)
    identifier_json = JSON(id_a)

    uri = "/repositories/12345/resources/import_#{SecureRandom.hex}"

    @records << JSONModel::JSONModel(:resource).
      from_hash({
                  :uri => uri,
                  :id_0 => id_a[0],
                  :id_1 => id_a[1],
                  :id_2 => id_a[2],
                  :id_3 => id_a[3],
                  :title => row['title'],
                  :level => 'collection',
                  :repository_processing_note => format_processing_note(row),
                  :extents => [format_extent(row, :portion => 'whole')].compact,
                  :dates => [format_date(row['date_expression'])].compact,
                  :rights_statements => [format_rights_statement(row)].compact,
                  :notes => [format_scopecontent_note(row)].compact,
                  :language => 'eng',
                })

  end


  def format_rights_statement(row)
    notes = []
    if row['access_conditions']
      notes << {
        :jsonmodel_type => "note_rights_statement",
        :label => "Access Conditions (eg Available for Reference. Not for Loan)",
        :type => 'additional_information',
        :content => [ row['access_conditions'] ]
      }
    end
    if row['use_conditions']
      notes << {
        :jsonmodel_type => "note_rights_statement",
        :label => "Use Conditions (eg copying not permitted)",
        :type => 'additional_information',
        :content => [ row['use_conditions'] ]
      }
    end

    {
      :rights_type => 'other',
      :other_rights_basis => 'donor',
      :start_date => Time.now.to_date.iso8601,
      :notes => notes
    }
  end


  def format_processing_note(row)
    [row['processing_note_1'], row['processing_note_2']].compact.join(' ')
  end


  def format_date(date_string)
    return if date_string.nil?

    {
      :date_type => date_string =~ /-/ ? 'inclusive' : 'single',
      :label => 'creation',
      :expression => date_string || "No date provided"
    }
  end


  def format_extent(row, opts = {})
    return unless row['extent_number'] && row['extent_type']

    {
      :portion => opts.fetch(:portion) { 'part' },
      :extent_type => row['extent_type'],
      :container_summary => row['extent_container_summary'],
      :number => row['extent_number'],
    }
  end


  def format_scopecontent_note(row)
    return unless row['scopecontent_note']
    {
      :jsonmodel_type => 'note_multipart',
      :type => 'scopecontent',
      :subnotes =>[{
                     :jsonmodel_type => 'note_text',
                     :content => row['scopecontent_note']
                   }]
    }
  end


  def row_values(row)
    (0...row.size).map {|i| row[i] ? row[i].to_s.strip : nil}
  end


  def format_record(row)

    record_hash = {
      :uri => "/repositories/12345/archival_objects/import_#{SecureRandom.hex}",
      :title => row['title'],
      :component_id => row['component_id'],
      :level => format_level(row['level']),
      :dates => [format_date(row['date'])].compact,
      :extents => [format_extent(row)].compact,
      :instances => [format_instance(row)].compact,
      :notes => [],
      :linked_agents => [format_agent(row)].compact,
      :resource => {
        :ref => @resource_uri
      },
    }

    if row['processinfo_note']
      record_hash[:notes] << {
        :jsonmodel_type => 'note_multipart',
        :type => 'processinfo',
        :subnotes =>[{
                       :jsonmodel_type => 'note_text',
                       :content => row['processinfo_note']
                     }]
      }
    end


    record_hash
  end

end
