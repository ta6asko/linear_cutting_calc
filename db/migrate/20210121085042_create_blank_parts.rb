class CreateBlankParts < ActiveRecord::Migration[6.0]
  def change
    create_table :blank_parts do |t|
      t.belongs_to :blank, foreign_key: { to_table: :blanks }
      t.belongs_to :part, foreign_key: { to_table: :parts }
      t.integer :quantity, null: false, default: 1

      t.timestamps null: false
    end
  end
end
