require_relative 'converters/arrearage_converter'

# Check config for obsolete carriers importer
class SpreadsheetImporterConfigException < Exception
end

raise SpreadsheetImporterConfigException.new("Please specify an agent uri for AppConfig[:obsolete_carriers_authorizer_agent_uri]") unless AppConfig.has_key?(:obsolete_carriers_authorizer_agent_uri)

#ref = JSONModel.parse_reference(AppConfig[:obsolete_carriers_authorizer_agent_uri])


invalid_uri_exception =
  SpreadsheetImporterConfigException.new("Invalid agent uri in " +
                                         "AppConfig[:obsolete_carriers_authorizer_agent_uri]: " +
                                         AppConfig[:obsolete_carriers_authorizer_agent_uri])

(leadingnil, agents, subagent, id) = AppConfig[:obsolete_carriers_authorizer_agent_uri].split('/')
unless agents == 'agents'
  raise invalid_uri_exception
end
model = case subagent
        when 'people'
          AgentPerson
        when 'corporate_entities'
          AgentCorporateEntity
        when 'families'
          AgentFamily
        when 'software'
          AgentSoftware
        else
          raise invalid_uri_exception
        end
  agent = model[id.to_i]
  if agent.nil?
    raise SpreadsheetImporterConfigException.new("Agent not found for uri specified in AppConfig[:obsolete_carriers_authorizer_agent_uri]: #{AppConfig[:obsolete_carriers_authorizer_agent_uri]}")
  end



# Work around small difference in rubyzip API
module Zip
  if !defined?(Error)
    class Error < StandardError
    end
  end
end

