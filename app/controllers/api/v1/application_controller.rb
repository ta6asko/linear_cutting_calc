# frozen_string_literal: true

module Api
  module V1
    class ApplicationController < ::ApplicationController
      def default_serializer_options
        {
          namespace: 'Api::V1',
          root: :data
        }
      end
    end
  end
end
