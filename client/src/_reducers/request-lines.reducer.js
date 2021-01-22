import { requestLineConstants } from '../_constants';

export function requestLines(state = {}, action) {
  switch (action.type) {
    case requestLineConstants.INDEX_REQUEST:
      return {
        loading: true
      }
    case requestLineConstants.INDEX_SUCCESS:
      return {
        requestLines: action.items.data
      }
    case requestLineConstants.INDEX_FAILURE:
      return {
        error: action.error
      }
    default:
      return state
  }
}
