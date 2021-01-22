import axios from 'axios'
import { authHeader } from '../_helpers'
import { merge } from 'lodash'
import { camelCaseKeys } from '../_utils/deepTransformKeys'

export const userService = {
  login,
  logout,
  register,
  fetchRequestedItems
}

function register(user) {
  const requestOptions = {
    baseURL: 'http://localhost:3001/api/v1',
    method: 'POST',
    url: '/auth',
    data: JSON.stringify(
      {
        email: user.email,
        password: user.password,
        password_confirmation: user.password_confirmation
      }
    ),
    headers: { 'Content-Type': 'application/json' }
  }

  return axios.request(requestOptions)
    .then((response) => {

      localStorage.setItem('user', JSON.stringify(camelCaseKeys(merge(response.data.data, response.headers))))

      return response
    })
    .catch((response) => {

      throw response
    })
}

function login(email, password) {
  const requestOptions = {
    baseURL: 'http://localhost:3001/api/v1',
    method: 'POST',
    url: '/auth/sign_in',
    data: JSON.stringify({ email, password }),
    headers: { 'Content-Type': 'application/json' }
  }

  return axios.request(requestOptions)
    .then((response) => {

      localStorage.setItem('user', JSON.stringify(camelCaseKeys(merge(response.data.data, response.headers))))

      return response
    })
    .catch((response) => {

      throw response
    })
}

function logout() {
  localStorage.removeItem('user')
}

function fetchRequestedItems(user) {
  const requestOptions = {
    baseURL: 'http://localhost:3001/api/v1',
    method: 'GET',
    url: `/l2on/users/${user.id}/user_items`,
    headers: authHeader()
  }

  return axios.request(requestOptions)
    .then((response) => {
      // localStorage.setItem('user', JSON.stringify(camelCaseKeys(merge(response.data.data, response.headers))))

      return response
    })
    .catch((response) => {

      throw response
    })
}

// function handleResponse(response) {
//     return response.text().then(text => {
//         const data = text && JSON.parse(text)
//         if (!response.ok) {
//             if (response.status === 401) {
//                 // auto logout if 401 response returned from api
//                 logout()
//                 // location.reload(true)
//             }

//             const error = (data && data.message) || response.statusText
//             return Promise.reject(error)
//         }

//         return data
//     })
// }
