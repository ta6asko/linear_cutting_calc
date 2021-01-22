# frozen_string_literal: true

Rails.application.routes.draw do
  namespace :api, defaults: { format: 'json' } do
    namespace :v1 do
      mount_devise_token_auth_for 'User', at: 'auth'

      resources :users, only: [] do
        resources :tasks, only: [:index, :create, :update, :destroy]
      end

      resources :tasks, only: [] do
        resources :blanks, only: [:index, :create, :update, :destroy]
        resources :parts, only: [:index, :create, :update, :destroy]
        resources :blank_parts, only: :index
      end
    end
  end
end
