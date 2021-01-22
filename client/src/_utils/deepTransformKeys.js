import { curry, pipe, T, cond, identity, fromPairs } from 'ramda'
import { isArray, camelCase, snakeCase, isObject, map } from 'lodash'

const _map = curry((f, x) => map(x, f))

const deepTransformKeys = curry((f, value) => cond([
  [isArray, _map(deepTransformKeys(f))],
  [isObject, pipe(_map((value, key) => [f(key), deepTransformKeys(f, value)]), fromPairs)],
  [T, identity],
])(value))

export const camelCaseKeys = (object) => deepTransformKeys(camelCase, object)
export const snakeCaseKeys = (object) => deepTransformKeys(snakeCase, object)


export default { camelCaseKeys, snakeCaseKeys }
