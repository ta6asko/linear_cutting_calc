class CreateParts < ActiveRecord::Migration[6.0]
  def change
    create_table :parts do |t|
      t.belongs_to :task, foreign_key: { to_table: :tasks }

      t.string :description
      t.integer :length, null: false
      t.integer :quantity, null: false, default: 1

      t.timestamps null: false
    end
  end
end
