# frozen_string_literal: true

module Api
  module V1
    class UserSerializer < ApplicationSerializer
      attributes :email
    end
  end
end
