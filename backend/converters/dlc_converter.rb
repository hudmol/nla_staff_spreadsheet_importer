class DLCConverter < Converter

  def self.instance_for(type, input_file)
    if type == "dlc"
      self.new(input_file)
    else
      nil
    end
  end


  def self.import_types(show_hidden = false)
    [
      {
        :name => "dlc",
        :description => "Digital Library Collections CSV"
      }
    ]
  end


  def self.profile
    "Convert a DLC CSV export to ArchivesSpace JSONModel records"
  end


  def initialize(input_file)
    super
    @batch = ASpaceImport::RecordBatch.new
    @input_file = input_file
    @records = []

    @columns = %w(level resource_id ud_int_2 container_type container_indicator component_id
                  title date extent_number extent_type extent_physical_details extent_dimensions
                  scopecontent_note creator processinfo_note)

    @level_map = {
      'Collection' => 'collection',
      'Class' => 'class',
      'Series' => 'series',
      'File' => 'file',
      'Item' => 'item'
    }

    @digital_object_uris = {}
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
          @resource_uri = get_or_create_resource(values_map)
          if @resource_uri.nil?
            raise "No resource defined"
          end

        when 'class'
          @class_uri = get_or_create_class(values_map)
          @series_uri = nil

        when 'series'
          @series_uri = get_or_create_series(values_map)

        when 'file'
          add_file(values_map)

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

  def get_or_create_resource(row)
    id_a = row['resource_id'].split(/\s+/)
    id_a = id_a + Array.new(4 - id_a.length)
    identifier_json = JSON(id_a)

    if (resource = Resource[:identifier => identifier_json])
      resource.uri
    else
      uri = "/repositories/12345/resources/import_#{SecureRandom.hex}"
      title = row['title']

      date = format_date(row['date'])

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
                    :linked_agents => [format_agent(row)].compact,
                    :user_defined => format_user_defined(row),
                    :language => 'eng',
                  })

      uri
    end
  end


  def get_or_create_class(row)
    return @class_uri unless format_level(row['level']) == 'class'

    class_hash = format_record(row)

    @records << JSONModel::JSONModel(:archival_object).from_hash(class_hash)

    class_hash[:uri]
  end


  def get_or_create_series(row)
    return @series_uri unless format_level(row['level']) == 'series'

    series_hash = format_record(row)

    series_hash[:parent] = { :ref => @class_uri } if @class_uri

    @records << JSONModel::JSONModel(:archival_object).from_hash(series_hash)

    series_hash[:uri]
  end


  def get_or_create_digital_object(row)
    id = row['container_indicator']
    if (digital_object = DigitalObject[:digital_object_id => id])
      digital_object.uri
    elsif @digital_object_uris[id]
      @digital_object_uris[id]
    else
      uri = "/repositories/12345/digital_objects/import_#{SecureRandom.hex}"

      do_hash = {
        :uri => uri,
        :digital_object_id => id,
        :title => row['title'],
        :user_defined => format_user_defined(row)
      }

      @records << JSONModel::JSONModel(:digital_object).from_hash(do_hash)
      @digital_object_uris[id] = uri

      uri
    end
  end


  def get_or_create_top_container(type, indicator)
    if @top_container_uris[type + indicator]
      @top_container_uris[type + indicator]
    else
      uri = "/repositories/12345/top_containers/import_#{SecureRandom.hex}"

      tc_hash = {
        :uri => uri,
        :type => type,
        :indicator => indicator
      }

      @records << JSONModel::JSONModel(:top_container).from_hash(tc_hash)
      @top_container_uris[type + indicator] = uri

      uri
    end
  end


  def get_or_create_agent(primary_name)
    if (name = NamePerson[:primary_name => primary_name])
      AgentPerson[name.agent_person_id].uri
    else
      uri = "/agents/people/import_#{SecureRandom.hex}"

      agent_hash = {
        :uri => uri,
        :names => [
                   {
                     :primary_name => primary_name,
                     :sort_name_auto_generate => true,
                     :name_order => 'inverted',
                     :source => 'local'
                   }
                  ]
      }

      @records << JSONModel::JSONModel(:agent_person).from_hash(agent_hash)

      uri
    end
  end


  def add_file(row)
    file_hash = format_record(row)

    file_hash[:parent] = { :ref => @class_uri } if @class_uri
    file_hash[:parent] = { :ref => @series_uri } if @series_uri

    @records << JSONModel::JSONModel(:archival_object).from_hash(file_hash)
  end


  def add_item(row)
    item_hash = format_record(row)

    item_hash[:parent] = { :ref => @class_uri } if @class_uri
    item_hash[:parent] = { :ref => @series_uri } if @series_uri

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
      :number => row['extent_number'],
      :physical_details => row['extent_physical_details'],
      :dimensions => row['extent_dimensions']
    }
  end


  def format_instance(row)
    return unless row['container_type'] && row['container_indicator']

    if row['container_type'] == 'Digital Item'
      {
        :instance_type => 'digital_object',
        :digital_object => {
          :ref => get_or_create_digital_object(row)
        }
      }
    else
      {
        :instance_type => 'mixed_materials',
        :sub_container => {
          :top_container => {
            :ref => get_or_create_top_container(row['container_type'], row['container_indicator'])
          }
        }
      }
    end
  end


  def format_agent(row)
    return unless row['creator']

    {
      :role => 'creator',
      :ref => get_or_create_agent(row['creator'])
    }
  end


  def format_user_defined(row)
    if row['ud_int_2']
      {
        :integer_2 => row['ud_int_2']
      }
    end
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
      :resource => {
        :ref => @resource_uri
      },
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
