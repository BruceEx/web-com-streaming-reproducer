/* global clearInterval, console, CustomFunctions, setInterval */

/**
 * OJS: TESTRTD = Increments a value once a second with OJS tag.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function testrtd(
  incrementBy: string,
  increment2: string,
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += 1;
    invocation.setResult(`OJS: ${result}`);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
