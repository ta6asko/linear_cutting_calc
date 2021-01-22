# frozen_string_literal: true

module Api
  module V1
    class BlanksController < ApplicationController
      def index
        @blanks = Task.first.blanks
        render json: @blanks
      end
    end
  end
end
