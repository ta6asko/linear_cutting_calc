import React from 'react'
import { Link } from 'react-router-dom'
import { connect } from 'react-redux'

import { userActions } from '../../_actions'

class SignUp extends React.Component {
  constructor(props) {
    super(props)

    this.state = {
      user: {
        email: '',
        password: ''
      },
      submitted: false
    }

    this.handleChange = this.handleChange.bind(this)
    this.handleSubmit = this.handleSubmit.bind(this)
  }

  handleChange(event) {
    const { name, value } = event.target
    const { user } = this.state
    this.setState({
      user: {
        ...user,
        [name]: value
      }
    })
  }

  handleSubmit(event) {
    event.preventDefault()

    this.setState({ submitted: true })
    const { user } = this.state
    if (user.email && user.password ) {
      this.props.register(user)
    }
  }

  render() {
    const { user, submitted } = this.state
    return (
      <form name="form" className="form-signup" onSubmit={this.handleSubmit}>
        <h1 className="h3 mb-3 font-weight-normal">Registration</h1>

        <div className={'form-group' + (submitted && !user.email ? ' has-error' : '')}>
          <label htmlFor="email">Email</label>
          <input type="text" className="form-control" name="email" value={user.email} onChange={this.handleChange} />
          {submitted && !user.email &&
              <div className="help-block">Email is required</div>
          }
        </div>

        <div className={'form-group' + (submitted && !user.password ? ' has-error' : '')}>
          <label htmlFor="password">Password</label>
          <input type="password" className="form-control" name="password" value={user.password} onChange={this.handleChange} />
          {submitted && !user.password &&
              <div className="help-block">Password is required</div>
          }
        </div>

        <div className="form-group">
          <button className="btn btn-primary">Register</button>
          <Link to="/login" className="btn btn-link">Cancel</Link>
        </div>
      </form>
    )
  }
}

function mapState(state) {
}

const actionCreators = {
  register: userActions.register
}

export default connect(mapState, actionCreators)(SignUp)
