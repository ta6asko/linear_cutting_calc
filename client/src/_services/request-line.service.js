import axios from 'axios'
import { authHeader } from '../_helpers'

export const requestLineService = {
  index
}

function index(item) {
  const requestOptions = {
    baseURL: 'http://localhost:3001/api/v1',
    method: 'GET',
    url: `/tasks/${1}/blanks`,
    headers: authHeader()
  }

  return axios.request(requestOptions)
    .then((response) => {

      return response;
    })
    .catch((response) => {

      throw response;
    })
}
