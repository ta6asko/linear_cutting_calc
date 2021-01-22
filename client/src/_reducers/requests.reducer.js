import { requestConstants } from '../_constants';

export function requests(state = {}, action) {
  switch (action.type) {
    case requestConstants.INDEX_REQUEST:
      return {
        loading: true
      }
    case requestConstants.INDEX_SUCCESS:
      return {
        items: action.items.data
      }
    case requestConstants.INDEX_FAILURE:
      return {
        error: action.error
      }
    default:
      return state
  }
}
