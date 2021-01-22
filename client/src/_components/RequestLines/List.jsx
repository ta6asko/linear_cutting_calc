import React from 'react'
import { connect } from 'react-redux'

import { requestLineActions } from '../../_actions'

class List extends React.Component {
  componentDidMount() {
    const { id } = this.props.match.params.id

    this.props.fetchRequestLines(id)
  }

  render() {
    const { requestLines } = this.props.requestLines

    return (
      <div>
        { requestLines &&
          <ol className="list-group">
            <li className="list-group-item list-group-item-secondary">
              <div className="row">
                <div className="col-md-3">
                  Персонаж
                </div>
                <div className="col-md-2">
                  Цена
                </div>
                <div className="col-md-2">
                  Модификация
                </div>
                <div className="col-md-3">
                  Замечен
                </div>
                <div className="col-md-2">
                  Город
                </div>
              </div>
            </li>
            { requestLines.map((line, index) =>
              <li className="list-group-item" key={ line.id }>
                <div className="row">
                  <div className="col-md-3">
                    { line.trader.name }
                  </div>
                  <div className="col-md-2">
                    { line.price }
                  </div>
                  <div className="col-md-2">
                    { line.enchant }
                  </div>
                  <div className="col-md-3">
                    { line.date }
                  </div>
                  <div className="col-md-2">
                    { line.town }
                  </div>
                </div>
              </li>
            )}
          </ol>
        }
      </div>
    )
  }
}

function mapState(state) {
  const { requestLines } = state
  return { requestLines }
}

const actionCreators = {
  fetchRequestLines: requestLineActions.fetchRequestLines
}

export default connect(mapState, actionCreators)(List)
