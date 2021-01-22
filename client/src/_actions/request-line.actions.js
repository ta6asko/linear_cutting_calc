import { requestLineConstants } from '../_constants';
import { requestLineService } from '../_services';

export const requestLineActions = {
  fetchRequestLines: index
}

function index(item) {
  return dispatch => {
    dispatch(request(item))

    requestLineService.index(item)
      .then(
        items => dispatch(success(items)),
        error => dispatch(failure(error.toString()))
      )
  }

  function request(item) { return { type: requestLineConstants.INDEX_REQUEST, item } }
  function success(items) { return { type: requestLineConstants.INDEX_SUCCESS, items } }
  function failure(error) { return { type: requestLineConstants.INDEX_FAILURE, error } }
}
