import React from 'react'
import { Link } from 'react-router-dom'
import { connect } from 'react-redux'

import { userActions } from '../../_actions'

class SignIn extends React.Component {
  constructor(props) {
    super(props)

    // reset login status
    this.props.logout()

    this.state = {
      email: '',
      password: '',
      submitted: false
    }

    this.handleChange = this.handleChange.bind(this)
    this.handleSubmit = this.handleSubmit.bind(this)
  }

  handleChange(event) {
    const { name, value } = event.target
    this.setState({ [name]: value })
  }

  handleSubmit(event) {
    event.preventDefault()

    this.setState({ submitted: true })
    const { email, password } = this.state
    if (email && password) {
      this.props.login(email, password)
    }
  }

  render() {
    const { email, password, submitted } = this.state
    return (
      <form name="form" className="form-signin" onSubmit={this.handleSubmit}>
        <h1 className="h3 mb-3 font-weight-normal">Login</h1>
        <div className={'form-group' + (submitted && !email ? ' has-error' : '')}>
          <label htmlFor="email">Email</label>
          <input type="text" className="form-control" name="email" value={email} onChange={this.handleChange} />
          {submitted && !email &&
            <div className="help-block">Email is required</div>
          }
        </div>
        <div className={'form-group' + (submitted && !password ? ' has-error' : '')}>
          <label htmlFor="password">Password</label>
          <input type="password" className="form-control" name="password" value={password} onChange={this.handleChange} />
          {submitted && !password &&
              <div className="help-block">Password is required</div>
          }
        </div>
        <div className="form-group">
          <button className="btn btn-primary">Login</button>
          <Link to="/register" className="btn btn-link">Register</Link>
        </div>
      </form>
    )
  }
}

function mapState(state) {
  const { loggingIn } = state.authentication
  return { loggingIn }
}

const actionCreators = {
  login: userActions.login,
  logout: userActions.logout
}

export default connect(mapState, actionCreators)(SignIn)
