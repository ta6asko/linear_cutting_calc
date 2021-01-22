# frozen_string_literal: true

module Api
  module V1
    class TasksController < ApplicationController
      def index
        @tasks = Task.all
        render json: @tasks
      end
    end
  end
end
