import { combineReducers } from 'redux'

import { authentication } from './authentication.reducer'
import { registration } from './registration.reducer'
import { users } from './users.reducer'
import { alert } from './alert.reducer'
import { requests } from './requests.reducer'
import { requestLines } from './request-lines.reducer'

const rootReducer = combineReducers({
  authentication,
  registration,
  users,
  requestLines,
  requests,
  alert
})

export default rootReducer
