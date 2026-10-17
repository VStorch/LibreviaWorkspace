import { useEffect } from 'react'

/** Closing the panel mid-answer stops the request: nobody is left to read it. */
export function useCancelOnUnmount(requestId: { readonly current: string | null }): void {
  useEffect(
    () => () => {
      if (requestId.current !== null) void window.api.ai.cancel({ requestId: requestId.current })
    },
    [requestId],
  )
}
