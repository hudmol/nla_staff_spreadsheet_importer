class ObsoleteCarriersConverter < Converter

  def self.instance_for(type, input_file)
    if type == "obsolete_carriers"
      self.new(input_file)
    else
      nil
    end
  end


  def self.import_types(show_hidden = false)
    [
      {
        :name => "obsolete_carriers",
        :description => "Obsolete Carriers CSV"
      }
    ]
  end


  def self.profile
    "Convert a CSV containing data on obsolete carrier collections to ArchivesSpace JSONModel records"
  end


  def initialize(input_file)
    super
    @batch = ASpaceImport::RecordBatch.new
    @input_file = input_file
    @records = []

    @columns = %w(
                  level
                  resource_id
                  component_id
                  repository_processing_note
                  instance_type
                  container_type
                  container_indicator
                  container_barcode
                  title
                  scopecontent_note
                  date_expression
                  extent_number
                  extent_type
                  extent_container_summary
                  extent_physical_details
                  extent_dimensions
                  subject_genre
                 )

    @level_map = {
      'Collection' => 'collection',
      'Item' => 'item'
    }

    @resource_uris = {}
    @top_container_uris = {}
  end


  def run
    rows = CSV.read(@input_file)

    begin
      while(row = rows.shift)
        values = row_values(row)

        next if values.compact.empty?

        values_map = Hash[@columns.zip(values)]

        case format_level(values_map['level'])

        when 'collection'
          get_or_create_resource(values_map)

        when 'item'
          add_item(values_map)

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

  def get_resource(id)
    get_or_create_resource({'resource_id' => id}, :create => false)
  end


  def get_or_create_resource(row, opts = {})

    create = opts.fetch(:create, true)

    if @resource_uris[row['resource_id']]
      @resource_uris[row['resource_id']]
    else
      id_a = [row['resource_id']]
      id_a = id_a + Array.new(4 - id_a.length)
      identifier_json = JSON(id_a)

      if (resource = Resource[:identifier => identifier_json])
        @resource_uris[row['resource_id']] = resource.uri
        resource.uri
      else
        raise "No resource for item #{row['component_id']} with Collection Number #{row['resource_id']}" unless create

        uri = "/repositories/12345/resources/import_#{SecureRandom.hex}"
        title = row['title']

        date = format_date(row['date_expression'])

        @records << JSONModel::JSONModel(:resource).from_hash({
                    :uri => uri,
                    :id_0 => id_a[0],
                    :id_1 => id_a[1],
                    :id_2 => id_a[2],
                    :id_3 => id_a[3],
                    :title => title,
                    :level => 'collection',
                    :extents => [format_extent(row, :portion => 'whole')].compact,
                    :dates => [date].compact,
                    :language => 'eng',
                  })

        @resource_uris[row['resource_id']] = uri
        uri
      end
    end
  end


  def get_or_create_top_container(row)
    tc_key = row['container_type'] + row['container_indicator']
    if @top_container_uris[tc_key]
      @top_container_uris[tc_key]
    else
      uri = "/repositories/12345/top_containers/import_#{SecureRandom.hex}"

      tc_hash = {
        :uri => uri,
        :type => row['container_type'],
        :indicator => row['container_indicator'],
      }

      @records << JSONModel::JSONModel(:top_container).from_hash(tc_hash)
      @top_container_uris[tc_key] = uri

      uri
    end
  end


  def add_item(row)
    item_hash = format_record(row)

    @records << JSONModel::JSONModel(:archival_object).from_hash(item_hash)
  end


  def format_level(level_string)
    @level_map[level_string]
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
      :physical_details => row['extent_physical_details'],
      :dimensions => row['extent_dimensions']
    }
  end


  def format_instance_type(type)
    return 'mixed_materials' unless type

    @instance_type_map ||= {
      'Accession' => 'accession',
      'Audio' => 'audio',
      'Computer Disks' => 'computer_disks',
      'Graphic Materials' => 'graphic_materials',
      'Mixed Materials' => 'mixed_materials',
      'Moving Images' => 'moving_images'
    }

    @instance_type_map.fetch(type, type)
  end


  def format_instance(row)
    return unless row['container_type'] && row['container_indicator']

    {
      :instance_type => format_instance_type(row['instance_type']),
      :sub_container => {
        :type_2 => 'object',
        :indicator_2 => row['container_barcode'],
        :top_container => {
          :ref => get_or_create_top_container(row)
        }
      }
    }
  end


  def get_subject_uri(term)
    @subject_uris ||= {}
    return @subject_uris[term] if @subject_uris.has_key?(term)

    subject_json = JSONModel::JSONModel(:subject).from_hash({
      :source => 'local',
      :vocabulary => '/vocabularies/1',
      :terms => [{
                   :vocabulary => '/vocabularies/1',
                   :term_type => 'genre_form',
                   :term => term
      }]
    })

    subject = Subject.find_matching(subject_json)

    raise "No subject found for '#{term}'" unless subject

    @subject_uris[term] = subject.uri
    @subject_uris[term]
  end


  def create_event(record_uri)
    uri = "/repositories/12345/events/import_#{SecureRandom.hex}"

    event_hash = {
      :uri => uri,
      :event_type => 'ingestion',
      :outcome => 'fail',
      :linked_agents => [{
                           :role => 'authorizer',
                           :ref => AppConfig[:obsolete_carriers_authorizer_agent_uri]
                         }],
      :linked_records => [{
                            :role => 'source',
                            :ref => record_uri
                          }],
      :date => {
        :begin => Time.now.strftime('%Y-%m-%d'),
        :date_type => 'single',
        :label => 'agent_relation'
      }
    }

    @records << JSONModel::JSONModel(:event).from_hash(event_hash)
    uri
  end


  def row_values(row)
    (0...row.size).map {|i| row[i] ? row[i].to_s.strip : nil}
  end


  def format_record(row)

    raise "No subject provided for '#{row['title']}' (#{row['component_id']})" unless row['subject_genre']
    
    uri = "/repositories/12345/archival_objects/import_#{SecureRandom.hex}"

    record_hash = {
      :uri => uri,
      :title => row['title'],
      :component_id => row['component_id'],
      :level => format_level(row['level']),
      :repository_processing_note => row['repository_processing_note'],
      :dates => [format_date(row['date'])].compact,
      :extents => [format_extent(row)].compact,
      :instances => [format_instance(row)].compact,
      :notes => [],
      :subjects => [{ :ref => get_subject_uri(row['subject_genre']) }],
      :linked_events => [{ :ref => create_event(uri) }],
      :resource => {
        :ref => get_resource(row['resource_id'])
      }
    }

    if row['scopecontent_note']
      record_hash[:notes] << {
        :jsonmodel_type => 'note_multipart',
        :type => 'scopecontent',
        :subnotes =>[{
                       :jsonmodel_type => 'note_text',
                       :content => row['scopecontent_note']
                     }]
      }
    end
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
