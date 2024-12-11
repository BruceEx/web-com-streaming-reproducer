/* global clearInterval, console, CustomFunctions, setInterval */

/**
 * OJS: TESTRTD = Increments a value once a second with OJS tag.
 * @customfunction
 * @volatile
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 */
export function testrtd(first, second)
{
    return Math.floor(Math.random() * first) + second;
}
