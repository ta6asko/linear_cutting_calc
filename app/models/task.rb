# frozen_string_literal: true

class Task < ApplicationRecord
  has_many :parts
  has_many :blanks
  has_many :blank_parts, through: :blanks
end
