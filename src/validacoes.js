// Data Validation Functions

/**
 * Validates if the input is a non-empty string.
 * @param {string} input - The input string to validate.
 * @returns {boolean} - True if valid, false otherwise.
 */
function validateNonEmptyString(input) {
    return typeof input === 'string' && input.trim() !== '';
}

/**
 * Validates if the input is a number.
 * @param {any} input - The input to validate.
 * @returns {boolean} - True if valid, false otherwise.
 */
function validateNumber(input) {
    return typeof input === 'number' && !isNaN(input);
}

/**
 * Validates if the input is within a specified range.
 * @param {number} input - The input number to validate.
 * @param {number} min - The minimum valid value.
 * @param {number} max - The maximum valid value.
 * @returns {boolean} - True if valid, false otherwise.
 */
function validateRange(input, min, max) {
    return validateNumber(input) && input >= min && input <= max;
}

module.exports = { validateNonEmptyString, validateNumber, validateRange };