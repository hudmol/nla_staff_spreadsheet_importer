require 'date'
require 'rubyXL'


# FIXME: just temporary stuff here.
if false
  class PermissiveEnumSource
    def method_missing(*)
      true
    end
  end

  JSONModel::init(:allow_other_unmapped => AppConfig[:allow_other_unmapped],
                  :enum_source => PermissiveEnumSource.new)
end


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

    # Now print the remaining data rows
    begin
      while (row = rows.next)
        jsonmodel = RowMapper.new(@headers.zip(row_values(row))).jsonmodel

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

    def initialize(row)
      @row = Hash[row]
    end

    def jsonmodel
      RequestContext.open(:current_username => "admin") do
        if @row['level'].downcase == 'collection'
          ResourceArrearage.from_row(@row).jsonmodel
        else
          ArchivalObjectArrearage.from_row(@row).jsonmodel
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


  class ParentHandler

    @uris = {}

    def self.parent_for(collection_id)
      if @uris[collection_id]
        return @uris[collection_id]
      end

      identifier = parse_identifier(collection_id)

      if (resource = Resource[:identifier => identifier])
        resource.uri
      end
    end


    def self.parse_identifier(collection_id)
      JSON((collection_id.split(/\//) + [nil, nil, nil, nil]).take(4))
    end


    def self.record_uri(collection_id, uri)
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

        name = NamePerson.filter(:primary_name => primary_name).filter(Sequel.expr(:rest_of_name).cast_string(CAST_TO_STRING) => rest_of_name).first

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
        name = NameCorporateEntity.filter(Sequel.expr(:primary_name).cast_string(CAST_TO_STRING) => primary_name).first

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

    attr_reader :row, :jsonmodel

    LEVELS = {'set' => 'file'}

    def initialize(row, jsonmodel)
      @row = row
      @jsonmodel = jsonmodel

      jsonmodel.level = LEVELS.fetch(row['level'].downcase, row['level'].downcase)
      jsonmodel.repository_processing_note = row['processing_note']

      jsonmodel.title = row['title']
      jsonmodel.instances = load_instances(row)
      jsonmodel.dates = load_dates(row)
      jsonmodel.extents = load_extents(row)
      jsonmodel.linked_agents = load_linked_agents(row)
    end


    protected

    def load_instances(row)
      return [] unless row['barcode']

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

      [{
         'jsonmodel_type' => 'instance',
         'instance_type' => 'graphic_materials',
         'container' => {
           'barcode_1' => row['barcode'],
           'indicator_2' => row['parent_container'],
           'container_locations' => container_locations,
         }
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

      [{
         'jsonmodel_type' => 'extent',
         'portion' => extent_portion,
         'number' => extent_number,
         'physical_details' => extent_physical_details,
         'extent_type' => 'photographic_prints',
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

    def self.from_row(row)
      new(row, ASpaceImport::JSONModel(:resource).new)
    end


    def initialize(row, jsonmodel)
      super

      if row['collection_id']
        row['collection_id'].split(/\//).each_with_index do |elt, i|
          jsonmodel["id_#{i}"] = elt
        end
      end

      if row['series_statement']
        jsonmodel['finding_aid_series_statement'] = row['series_statement']
      end

      if row['catalogued_note']
        jsonmodel['collection_management'] = {
          'jsonmodel_type' => 'collection_management',
          'cataloged_note' => row['catalogued_note']
        }
      end

      import_uri = "/repositories/12345/resources/import_#{SecureRandom.hex}"

      ParentHandler.record_uri(row['collection_id'], import_uri)

      jsonmodel['uri'] = import_uri
    end

  end


  class ArchivalObjectArrearage < Arrearage

    def self.from_row(row)
      new(row, ASpaceImport::JSONModel(:archival_object).new)
    end


    def initialize(row, jsonmodel)
      super

      if row['component_id']
        jsonmodel.component_id = row['component_id']
      end

      collection_id = row.fetch('collection_id') {
        raise "Missing value for collection ID"
      }

      jsonmodel.resource = {'ref' => ParentHandler.parent_for(collection_id)}
    end

  end

end



# ArrearageConverter.new("/home/mst/projects/nla-archivesspace/specs/Arrearage Template.xlsx").run
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
