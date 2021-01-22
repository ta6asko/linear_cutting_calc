import React from 'react'
import { Router, Route, Switch, Redirect } from 'react-router-dom'
import { connect } from 'react-redux'

import { history } from './_helpers'
import { alertActions } from './_actions'
import { PrivateRoute } from './_components/PrivateRoute'
import SignUp from './_components/User/SignUp'
import SignIn from './_components/User/SignIn'
import RequestLineList from './_components/RequestLines/List'
import Dashboard from './_components/Dashboard'
import Header from './_components/Header'

class App extends React.Component {
  constructor(props) {
    super(props)

    history.listen((location, action) => {
      this.props.clearAlerts()
    })
  }

  render() {
    const { alert } = this.props

    return (
      <div className="container">
        {alert.message &&
          <div className={`alert ${alert.type}`}>{alert.message}</div>
        }
        <Router history={history}>
          <Header/>
          <Switch>
            <div className="container mt-5">
              {/*<PrivateRoute exact path="/" component={Dashboard} />*/}
              {/*<PrivateRoute exact path="/requests/:id" component={RequestLineList} />*/}
              <Route exact path="/" component={Dashboard} />
              <Route exact path="/requests/:id" component={RequestLineList} />
              <Route path="/login" component={SignIn} />
              <Route path="/register" component={SignUp} />
              <Redirect from="*" to="/" />
            </div>
          </Switch>
        </Router>
      </div>
    )
  }
}

function mapState(state) {
  const { alert } = state
  return { alert }
}

const actionCreators = {
  clearAlerts: alertActions.clear
}

export default connect(mapState, actionCreators)(App)
