# frozen_string_literal: true

class Part < ApplicationRecord
  belongs_to :task
end
