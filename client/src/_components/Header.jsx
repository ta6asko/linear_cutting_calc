import React from 'react'
import { Link } from 'react-router-dom'
import { connect } from 'react-redux'

class Header extends React.Component {
  render() {
    const { user } = this.props
    return (
      <nav className="navbar navbar-expand-lg navbar-light bg-light">
        <a className="navbar-brand" href="/">Избраное</a>
        <button className="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavAltMarkup" aria-controls="navbarNavAltMarkup" aria-expanded="false" aria-label="Toggle navigation">
          <span className="navbar-toggler-icon"></span>
        </button>
        <div className="collapse navbar-collapse" id="navbarNavAltMarkup">
          <div className="navbar-nav">
            { !user &&
              <Link to="/login" className="nav-item nav-link active">Войти</Link>
            }
            { !user &&
              <Link to="/register" className="nav-item nav-link">Регистрация</Link>
            }
            { user &&
              <Link to="/login" className="nav-item nav-link">Выйти</Link>
            }
          </div>
        </div>
      </nav>
    )
  }
}

function mapState(state) {
  const { authentication } = state
  const { user } = authentication
  return { user }
}

const actionCreators = {
}

export default connect(mapState, actionCreators)(Header)
