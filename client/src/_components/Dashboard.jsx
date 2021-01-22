import React from 'react'
import { connect } from 'react-redux'

import RequestsList from './Requests/List'

class Dashboard extends React.Component {
  render() {
    return (
      <div className="mt-5">
       {/* <h6 className="border-bottom border-gray pb-2 mb-0">
          { requests.loading && <em>Loading requests...</em> }
          { requests.error && <span className="text-danger">ERROR: { items.error }</span> }
        </h6>*/}

        <RequestsList/>
      </div>
    )
  }
}

function mapState(state) {
  return {}
}

const actionCreators = {
}

export default connect(mapState, actionCreators)(Dashboard)
