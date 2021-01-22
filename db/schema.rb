# This file is auto-generated from the current state of the database. Instead
# of editing this file, please use the migrations feature of Active Record to
# incrementally modify your database, and then regenerate this schema definition.
#
# This file is the source Rails uses to define your schema when running `rails
# db:schema:load`. When creating a new database, `rails db:schema:load` tends to
# be faster and is potentially less error prone than running all of your
# migrations from scratch. Old migrations may fail to apply correctly if those
# migrations use external dependencies or application code.
#
# It's strongly recommended that you check this file into your version control system.

ActiveRecord::Schema.define(version: 2021_01_21_085042) do

  # These are extensions that must be enabled in order to support this database
  enable_extension "plpgsql"

  create_table "blank_parts", force: :cascade do |t|
    t.bigint "blank_id"
    t.bigint "part_id"
    t.integer "quantity", default: 1, null: false
    t.datetime "created_at", precision: 6, null: false
    t.datetime "updated_at", precision: 6, null: false
    t.index ["blank_id"], name: "index_blank_parts_on_blank_id"
    t.index ["part_id"], name: "index_blank_parts_on_part_id"
  end

  create_table "blanks", force: :cascade do |t|
    t.bigint "task_id"
    t.string "description"
    t.integer "length", null: false
    t.integer "quantity", default: 1, null: false
    t.datetime "created_at", precision: 6, null: false
    t.datetime "updated_at", precision: 6, null: false
    t.index ["task_id"], name: "index_blanks_on_task_id"
  end

  create_table "parts", force: :cascade do |t|
    t.bigint "task_id"
    t.string "description"
    t.integer "length", null: false
    t.integer "quantity", default: 1, null: false
    t.datetime "created_at", precision: 6, null: false
    t.datetime "updated_at", precision: 6, null: false
    t.index ["task_id"], name: "index_parts_on_task_id"
  end

  create_table "tasks", force: :cascade do |t|
    t.bigint "user_id"
    t.decimal "cutting_thickness", precision: 15, scale: 2, null: false
    t.string "description"
    t.datetime "created_at", precision: 6, null: false
    t.datetime "updated_at", precision: 6, null: false
    t.index ["user_id"], name: "index_tasks_on_user_id"
  end

  create_table "users", force: :cascade do |t|
    t.datetime "created_at", precision: 6, null: false
    t.datetime "updated_at", precision: 6, null: false
  end

  add_foreign_key "blank_parts", "blanks"
  add_foreign_key "blank_parts", "parts"
  add_foreign_key "blanks", "tasks"
  add_foreign_key "parts", "tasks"
  add_foreign_key "tasks", "users"
end
