# frozen_string_literal: true

class BlankParts < ApplicationRecord
  belongs_to :part
  belongs_to :blank
end
