require 'date'
require 'rubyXL'
require 'set'


class ArrearageConverter < Converter

  START_MARKER = /ArchivesSpace field code/

  def self.instance_for(type, input_file)
    if type == "arrearage"
      self.new(input_file)
    else
      nil
    end
  end


  def self.import_types(show_hidden = false)
    [
      {
        :name => "arrearage",
        :description => "Import Arrearage spreadsheet"
      }
    ]
  end


  def self.profile
    "Convert an Arrearage spreadsheet to ArchivesSpace JSONModel records"
  end


  def initialize(input_file)
    super
    @batch = ASpaceImport::RecordBatch.new
    @input_file = input_file
  end


  def run
    workbook = RubyXL::Parser.parse(@input_file)
    sheet = workbook[0]

    rows = sheet.enum_for(:each)

    while @headers.nil? && (row = rows.next)
      if row[0] && row[0].value =~ START_MARKER
        @headers = row_values(row)

        # Skip the human readable header too
        rows.next
      end
    end

    context = {
      :collection_handler => CollectionHandler.new,
      :parent_handler => ParentHandler.new,
      :position => 0,
    }

    # Now print the remaining data rows
    begin
      while (row = rows.next)
        values = row_values(row)

        next if values.compact.empty?

        jsonmodel = RowMapper.new(@headers.zip(values)).jsonmodel(context)

        p jsonmodel

        @batch << jsonmodel
      end
    rescue StopIteration
    end
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

  def row_values(row)
    (1...row.size).map {|i| (row[i] && row[i].value) ? row[i].value.to_s.strip : nil}
  end


  class RowMapper

    class RowWrapper

      def initialize(row)
        @row = row
        @accessed_properties = Set.new
      end

      def method_missing(method, *args)
        if method == :fetch || method == :[]
          @accessed_properties << args.first
        end

        @row.send(method, *args)
      end

      def warn_of_missed_properties
        (@row.keys - @accessed_properties).each do |missed_property|
          $stderr.puts("Didn't use property: #{missed_property}")
        end
      end

    end


    def initialize(row)
      @row = RowWrapper.new(Hash[row])
    end

    def jsonmodel(context)
      RequestContext.open(:current_username => "admin") do
        if @row['level'].downcase =~ /collection/
          ResourceArrearage.from_row(@row, context).jsonmodel
        else
          ArchivalObjectArrearage.from_row(@row, context).jsonmodel
        end
      end
    end

  end


  class LocationHandler

    def self.get_or_create(location)
      aspace_location = {:building => "The National Library of Australia",
                         :room => location[:room],
                         :coordinate_1_label => "Row/drawer",
                         :coordinate_1_indicator => location[:row],
                         :coordinate_2_label => "Unit",
                         :coordinate_2_indicator => location[:unit],
                         :coordinate_3_label => "Shelf",
                         :coordinate_3_indicator => location[:shelf]}

      if (location = Location[aspace_location])
        location.uri
      else
        Location.create_from_json(JSONModel::JSONModel(:location).from_hash(aspace_location)).uri
      end
    end

  end


  class CollectionHandler

    def initialize
      @uris = {}
    end

    def parent_for(collection_id)
      if @uris[collection_id]
        return @uris[collection_id]
      end

      identifier = parse_identifier(collection_id)

      if (resource = Resource[:identifier => identifier])
        resource.uri
      else
        raise "No parent found matching collection identifier: #{collection_id}"
      end
    end


    def parse_identifier(collection_id)
      JSON((collection_id.split(/\//) + [nil, nil, nil, nil]).take(4))
    end


    def record_uri(collection_id, uri)
      @uris[collection_id] = uri
    end

  end


  class AgentHandler

    # Working around the fact that agent names are using CLOBs here.  Sigh.
    # With any luck the name tables won't get big enough that a full table scan
    # matters much.  If it does, we'll need to modify ArchivesSpace to give us
    # an indexed column to hit when looking up by name.

    CAST_TO_STRING = 'varchar(2048)'

    def self.get_or_create(primary_name, rest_of_name, agent_type)
      agent = nil

      if agent_type == 'agent_person'

        name = if ASpaceEnvironment.demo_db?
                 NamePerson.filter(:primary_name => primary_name).filter(Sequel.expr(:rest_of_name).cast_string(CAST_TO_STRING) => rest_of_name).first
               else
                 NamePerson.filter(:primary_name => primary_name).filter(:rest_of_name => rest_of_name).first
               end

        if name
          agent = AgentPerson[name.agent_person_id]
        else
          agent = AgentPerson.create_from_json(JSONModel::JSONModel(:agent_person).from_hash('names' => [
                                                                                               {
                                                                                                 'jsonmodel_type' => 'name_person',
                                                                                                 'primary_name' => primary_name,
                                                                                                 'rest_of_name' => rest_of_name,
                                                                                                 'name_order' => 'inverted',
                                                                                                 'source' => 'local',
                                                                                                 'sort_name_auto_generate' => true,
                                                                                               }
                                                                                             ]))
        end
      elsif agent_type == 'agent_corporate_entity'

        name = if ASpaceEnvironment.demo_db?
                 name = NameCorporateEntity.filter(Sequel.expr(:primary_name).cast_string(CAST_TO_STRING) => primary_name).first
               else
                 name = NameCorporateEntity.filter(:primary_name => primary_name).first
               end

        if name
          agent = AgentCorporateEntity[name.agent_corporate_entity_id]
        else
          agent = AgentCorporateEntity.create_from_json(JSONModel::JSONModel(:agent_corporate_entity).from_hash('names' => [
                                                                                                                  {
                                                                                                                    'jsonmodel_type' => 'name_corporate_entity',
                                                                                                                    'primary_name' => primary_name,
                                                                                                                    'source' => 'local',
                                                                                                                    'sort_name_auto_generate' => true,
                                                                                                                  }
                                                                                                                ]))
        end
      else
        raise "Unrecognised agent type: #{agent_type}"
      end


      agent.uri
    end

  end


  class Arrearage

    attr_reader :row, :jsonmodel, :context

    LEVELS = {
      'set' => 'file',
      'collection (multi)' => 'collection',
      'collection (single)' => 'collection',
    }

    def initialize(row, jsonmodel, context)
      @row = row
      @jsonmodel = jsonmodel
      @context = context

      jsonmodel.level = LEVELS.fetch(row['level'].downcase, row['level'].downcase)
      jsonmodel.repository_processing_note = if row['processing_note']
                                               "Previous location: " + row['processing_note']
                                             end

      jsonmodel.title = row['title']
      jsonmodel.instances = load_instances(row)
      jsonmodel.dates = load_dates(row)
      jsonmodel.extents = load_extents(row)
      jsonmodel.linked_agents = load_linked_agents(row)
    end


    protected

    def load_instances(row)

      container_locations = []

      if row['location_room']
        container_locations << {
          'start_date' => Date.today.strftime('%Y-%m-%d'),
          'ref' => ::ArrearageConverter::LocationHandler.get_or_create(:room => row['location_room'],
                                                                       :row => row['location_row'],
                                                                       :unit => row['location_unit'],
                                                                       :shelf => row['location_shelf']),
          'status' => 'current'
        }
      end

      container = {
        'type_1' => 'box',
        'container_locations' => container_locations,
      }

      if row['indicator_1']
        container['indicator_1'] = row['indicator_1']
      else
        # We need something...
        container['barcode_1'] = SecureRandom.hex
        container['indicator_1'] = 'Unknown'
      end


      if row['indicator_2']
        container['type_2'] = 'piece'
        container['indicator_2'] = row['indicator_2']
      end


      [{
         'jsonmodel_type' => 'instance',
         'instance_type' => 'graphic_materials',
         'container' => container
       }]
    end


    def load_dates(row)
      return [] unless row['date']

      [{
         'jsonmodel_type' => 'date',
         'date_type' => 'single',
         'expression' => row['date'],
         'label' => 'creation',
       }]
    end


    def load_extents(row)
      extent_portion = 'whole'
      extent_number = row.fetch('extent_number', '1')
      extent_dimensions = row.fetch('extent_dimensions', nil)
      extent_physical_details = row.fetch('extent_physical_details', nil)
      extent_type = row.fetch('extent_type', nil)

      [{
         'jsonmodel_type' => 'extent',
         'portion' => extent_portion,
         'number' => extent_number,
         'physical_details' => extent_physical_details,
         'extent_type' => extent_type,
         'dimensions' => extent_dimensions,
       }]
    end


    def load_linked_agents(row)

      result = []

      p row

      creator_count = row.keys.grep(/creator_[0-9]+_/).length

      (1..creator_count).each do |i|
        if row["creator_#{i}_primary_name"]
          result << {
            'ref' => AgentHandler.get_or_create(row["creator_#{i}_primary_name"], row["creator_#{i}_rest_of_name"], 'agent_person'),
            'role' => 'creator'
          }
        end
      end

      row.keys.grep(/linked_corporate_entity_agent(_[0-9]+\z)?/).each do |corporate_column|
        if row[corporate_column]
          result << {
            'ref' => AgentHandler.get_or_create(row[corporate_column], nil, 'agent_corporate_entity'),
            'role' => 'creator'
          }
        end
      end

      result
    end

  end



  class ResourceArrearage < Arrearage

    def self.from_row(row, context)
      new(row, ASpaceImport::JSONModel(:resource).new, context)
    end


    def initialize(row, jsonmodel, context)
      super

      if row['collection_id']
        row['collection_id'].split(/\//).each_with_index do |elt, i|
          jsonmodel["id_#{i}"] = elt
        end
      end

      if row['series_statement']
        jsonmodel['finding_aid_series_statement'] = row['series_statement']
      end

      jsonmodel.user_defined = load_user_defined_fields(row)

      import_uri = "/repositories/12345/resources/import_#{SecureRandom.hex}"

      context.fetch(:collection_handler).record_uri(row['collection_id'], import_uri)

      jsonmodel['uri'] = import_uri
      jsonmodel['position'] = (context[:position] += 1)
    end


    def load_user_defined_fields(row)
      udf = {}

      if row['bib_collection_level']
        udf['integer_2'] = row['bib_collection_level']
      end

      if row['pres_work_req']
        if row['pres_work_req'] =~ /y/i
          udf['enum_2'] = 'Preservation Required'
        elsif row['pres_work_req'] =~ /n/i
          udf['enum_2'] = 'Preservation Not Required'
        end
      end

      if row['digitisation_notes']
        udf['text_5'] = row['digitisation_notes']
      end


      if udf.empty?
        nil
      else
        {'jsonmodel_type' => 'user_defined'}.merge(udf)
      end
    end

  end


  class ParentHandler

    def set_uri(row, uri)
      @current_hierarchy ||= {}

      collection_hierarchy = parse_collection_hierarchy(row)

      # Chop the hierarchy back to the level of the current record
      @current_hierarchy = Hash[@current_hierarchy.map {|k, v|
                                  if k < collection_hierarchy
                                    [k, v]
                                  end
                                }.compact]

      # Record the URI of the current record
      @current_hierarchy[collection_hierarchy] = uri
    end


    def parent_for(row)
      collection_hierarchy = parse_collection_hierarchy(row)

      # Level 1 will be a resource record, Level 2 is a top-level AO (no
      # parent), so Level 3 is the first one we care about.
      if collection_hierarchy > 2
        parent_level = collection_hierarchy - 1

        @current_hierarchy.fetch(parent_level)
      else
        nil
      end
    end


    private

    def parse_collection_hierarchy(row)
      begin
        Integer(row['collection_hierarchy'])
      rescue ArgumentError
        raise ArgumentError.new("Collection hierarchy was not a number: #{row['collection_hierarchy']}")
      end
    end

  end


  class ArchivalObjectArrearage < Arrearage

    def self.from_row(row, context)
      new(row, ASpaceImport::JSONModel(:archival_object).new, context)
    end


    def initialize(row, jsonmodel, context)
      super

      import_uri = "/repositories/12345/archival_objects/import_#{SecureRandom.hex}"
      jsonmodel['uri'] = import_uri
      jsonmodel['position'] = (context[:position] += 1)


      component_id = [row['component_id'], row['item_id']].compact.join("")

      if !component_id.empty?
        jsonmodel.component_id = component_id
      end

      collection_id = row.fetch('collection_id') {
        raise "Missing value for collection ID"
      }

      jsonmodel.resource = {'ref' => context.fetch(:collection_handler).parent_for(collection_id)}

      context.fetch(:parent_handler).set_uri(row, import_uri)

      jsonmodel.parent = {'ref' => context.fetch(:parent_handler).parent_for(row)}
    end

  end

end



#
# Questions:
#
# graphic_materials OK for instance type?
#
# start date for location?  Set to today?
#
# default for extent_dimensions, extent_portion, extent_number?
#
# photographic_prints for extent type OK?
