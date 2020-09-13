Rails.application.routes.draw do
  root 'companies#index'
  resources :clients
  resources :contractors
  resources :partner_companies
  resources :employees do
    collection do
      get 'upload'
      post 'upload'
    end
  end
  resources :companies
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
