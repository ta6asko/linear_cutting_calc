# frozen_string_literal: true

module Api
  module V1
    class ApplicationSerializer < ActiveModel::Serializer
      attributes :id,
                 :created_at,
                 :updated_at

      def created_at
        object.created_at.in_time_zone.iso8601
      end

      def updated_at
        object.updated_at.in_time_zone.iso8601
      end
    end
  end
end
