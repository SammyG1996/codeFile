// src/extensions/customFormApp/Utils/getFetchApi.ts

export type FetchAPInfo = {
  spUrl: string,
  method: string,
  headers: any
};

export async function getFetchAPI(props: FetchAPInfo): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
  let attempt = 0;
  const maxRetries = 5;
  const baseDelay = 500;

  while (attempt < maxRetries) {
    try {
      const response = await fetch(props.spUrl, {
        method: props.method,
        headers: props.headers
      });

      // Check for any Rate limit headers and then back off
      const rateLimitRemaining = response.headers.get('RateLimit-Remaining');
      const rateLimitReset = response.headers.get('RateLimit-Reset');

      if (rateLimitRemaining && parseInt(rateLimitRemaining, 10) < 10 && rateLimitReset) {
        const restTime = parseInt(rateLimitReset, 10) * 1000;
        console.warn(`Approaching SharePoint API rate limit. Pausing for ${restTime}`);
        await new Promise((resolve) => setTimeout(resolve, restTime));
      } else if (!response.ok) {
        if (response.status === 429 || response.status === 503) {
          const retryAfterHeader = response.headers.get('Retry-After');
          const delay =
            (retryAfterHeader ? parseInt(retryAfterHeader, 10) * 1000 : 0) +
            baseDelay * (2 ** attempt);
          console.warn(`429 or 503 error received, Retrying in ${delay}ms. Attempt: ${attempt}`);
          await new Promise((resolve) => setTimeout(resolve, delay));
          attempt++;
        } else {
          return Promise.reject(`Fetch API request failed with staus ${response.status}: ${response.statusText}`);
        }
      } else {
        // ADDED: safe read of "X-HTTP-Method" without upsetting TS
        const headersObj = (props.headers ?? {}) as Record<string, unknown>;
        const xHttpMethod =
        typeof headersObj['X-HTTP-Method'] === 'string'
            ? (headersObj['X-HTTP-Method'] as string)
            : '';

        // DELETE if real method is DELETE, or X-HTTP-Method spoof is DELETE
        const isDelete =
        props.method?.toUpperCase() === 'DELETE' || xHttpMethod.toUpperCase() === 'DELETE';

        if (isDelete) {
          // Many SP delete endpoints return 200/204 with no content.
          // Returning text() (possibly "") avoids the old json() crash
          // while leaving all non-DELETE behavior untouched.
          const txt = await response.text();
          return Promise.resolve(txt);
        }

        const contentType = response.headers.get('Content-Type');
        if (contentType === 'text/xml') {
          return Promise.resolve(await response.text());
        } else {
          const jsonData = await response.json();
          if (Object.hasOwn(jsonData, 'd')) {
            return Promise.resolve(jsonData.d);
          } else {
            return Promise.resolve(jsonData);
          }
        }
      }
    } catch (error) {
      console.error(`Error during Fetch API request (attempt ${attempt + 1}): `, error);
      if (attempt >= maxRetries - 1) {
        console.error(`Fetch API request failed after ${maxRetries} attempts.`);
        return Promise.reject(`Fetch API request failed after ${maxRetries} attempts.`);
      }
      attempt++;
    }
  }
}