class CreateProjects < ActiveRecord::Migration[6.0]
  def change
    create_table :projects do |t|
      t.belongs_to :user, foreign_key: { to_table: :users }

      t.decimal :cutting_thickness, precision: 15, scale: 2, null: false
      t.string :description

      t.timestamps null: false
    end
  end
end
