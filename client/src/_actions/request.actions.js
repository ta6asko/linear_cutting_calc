import { requestConstants } from '../_constants';
import { requestService } from '../_services';

export const requestActions = {
  fetchRequests: index
}

function index(user) {
  return dispatch => {
    dispatch(request(user))

    requestService.index(user)
      .then(
        items => dispatch(success(items)),
        error => dispatch(failure(error.toString()))
      )
  }

  function request(item) { return { type: requestConstants.INDEX_REQUEST, item } }
  function success(items) { return { type: requestConstants.INDEX_SUCCESS, items } }
  function failure(error) { return { type: requestConstants.INDEX_FAILURE, error } }
}
