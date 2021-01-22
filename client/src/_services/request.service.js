import axios from 'axios'
import { authHeader } from '../_helpers'

export const requestService = {
  index
}

function index(user) {
  const requestOptions = {
    baseURL: 'http://localhost:3001/api/v1',
    method: 'GET',
    url: `/users/${1}/tasks`,
    headers: authHeader()
  }

  return axios.request(requestOptions)
    .then((response) => {
      console.log('SUC', response)
      return response
    })
    .catch((response) => {
      console.log('ERROR', response)

      throw response
    })
}
