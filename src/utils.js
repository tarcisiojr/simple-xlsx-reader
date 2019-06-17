const peek = (arr) => arr.length ? arr[arr.length - 1] : null;

const chainPromises = (promises) => {
    return promises.reduce((prev, task) => {
        return prev.then(() => task())
    }, Promise.resolve())
}

module.exports = {
    peek,
    chainPromises
}