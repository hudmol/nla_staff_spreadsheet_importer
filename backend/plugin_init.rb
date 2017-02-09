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


require_relative 'converters/donation_converter'

if ASConstants.VERSION =~ /1\.2\.0/

  #
  # PATCH TO RUN CONVERTER INSIDE THE REQUEST CONTEXT
  #
  # See https://github.com/archivesspace/archivesspace/pull/209
  #
  BatchImportRunner.class_eval do
    def run
      ticker = Ticker.new(@job)

      last_error = nil
      batch = nil
      success = false

      filenames = @json.job['filenames'] || []

      # Wrap the import in a transaction if the DB supports MVCC
      begin
        DB.open(DB.supports_mvcc?,
                :retry_on_optimistic_locking_fail => true) do

          begin
            @job.job_files.each_with_index do |input_file, i|
              ticker.log(("=" * 50) + "\n#{filenames[i]}\n" + ("=" * 50)) if filenames[i]
              converter = Converter.for(@json.job['import_type'], input_file.file_path)
              begin
                RequestContext.open(:create_enums => true,
                                    :current_username => @job.owner.username,
                                    :repo_id => @job.repo_id) do

                  converter.run

                  File.open(converter.get_output_path, "r") do |fh|
                    batch = StreamingImport.new(fh, ticker, @import_canceled)
                    batch.process
                    log_created_uris(batch)
                    success = true
                  end
                end
              ensure
                converter.remove_files
              end
            end
          rescue ImportCanceled
            raise Sequel::Rollback
          rescue JSONModel::ValidationException, ImportException, Converter::ConverterMappingError, Sequel::ValidationFailed, ReferenceError => e
            # Note: we deliberately don't catch Sequel::DatabaseError here.  The
            # outer call to DB.open will catch that exception and retry the
            # import for us.
            last_error = e

            # Roll back the transaction (if there is one)
            raise Sequel::Rollback
          end
        end
      rescue
        last_error = $!
      end

      if last_error

        ticker.log("\n\n" )
        ticker.log( "!" * 50 )
        ticker.log( "IMPORT ERROR" )
        ticker.log( "!" * 50 )
        ticker.log("\n\n" )

        if  last_error.respond_to?(:errors)

          ticker.log("#{last_error}") if last_error.errors.empty? # just spit it out if there's not explicit errors

          ticker.log("The following errors were found:\n")

          last_error.errors.each_pair { |k,v| ticker.log("\t#{k.to_s} : #{v.join(' -- ')}" ) }

          if last_error.is_a?(Sequel::ValidationFailed)
            ticker.log("\n" )
            ticker.log("%" * 50 )
            ticker.log("\n Full Error Message:\n #{last_error.to_s}\n\n")
          end

          if ( last_error.respond_to?(:invalid_object) && last_error.invalid_object )
            ticker.log("\n\n For #{ last_error.invalid_object.class }: \n #{ last_error.invalid_object.inspect  }")
          end

          if ( last_error.respond_to?(:import_context) && last_error.import_context )
            ticker.log("\n\nIn : \n #{ CGI.escapeHTML( last_error.import_context ) } ")
            ticker.log("\n\n")
          end
        else
          ticker.log("Error: #{CGI.escapeHTML(  last_error.inspect )}")
        end
        ticker.log("!" * 50 )
        raise last_error
      end
    end
  end

end
