import React from 'react'
import { connect } from 'react-redux'
import { Link } from 'react-router-dom'

import { requestActions } from '../../_actions'

class List extends React.Component {
  componentDidMount() {
    this.props.fetchRequests(this.props.user)
  }

  render() {
    const { requests } = this.props

    return (
      <div>
        { requests.items &&
          <div class="list-group">
            { requests.items.map((request, index) =>
              <Link to={`/requests/${request.id}`}
                    className="list-group-item list-group-item-action"
                    key={ request.id }>
                { request.id }
              </Link>
            )}
          </div>
        }
      </div>
    )
  }
}

function mapState(state) {
  const { authentication, requests } = state
  const { user } = authentication
  return { user, requests }
}

const actionCreators = {
  fetchRequests: requestActions.fetchRequests
}

export default connect(mapState, actionCreators)(List)
