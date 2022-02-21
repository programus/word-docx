module.exports = {
  /**
   * convert object to css text
   * @param {Object} obj 
   * @param {Boolean} important 
   * @returns 
   */
  object2CssText(obj, important) {
    return Object.entries(obj).map(e => `${e[0]}: ${e[1]}${important ? ' !important' : ''};`).join(' ')
  },

  /**
   * Add two css length units
   * @param {string} oldValue 
   * @param {string} newValue 
   * @returns {string} calc expression
   */
  addCssUnits(oldValue, newValue) {
    return oldValue ? `calc(${oldValue.replace(/(?:calc\s*\()?([^)]+)(?:\))?/g, '$1')} + ${newValue})` : newValue
  },
}