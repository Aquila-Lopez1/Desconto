// Utility functions and helpers

// Example function: Returns the maximum of two numbers
function max(a, b) {
    return a > b ? a : b;
}

// Example function: Returns the minimum of two numbers
function min(a, b) {
    return a < b ? a : b;
}

// Example function: Checks if a number is even
function isEven(num) {
    return num % 2 === 0;
}

module.exports = { max, min, isEven };