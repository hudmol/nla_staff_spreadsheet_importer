require_relative 'converters/arrearage_converter'


# Work around small difference in rubyzip API
module Zip
  if !defined?(Error)
    class Error < StandardError
    end
  end
end
