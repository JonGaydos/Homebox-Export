"""HTTP retry decision, adapted from skydio-transfer."""


def retry_policy(exception, status_code, retry_after, attempt, max_retries):
    """Decide whether to retry an HTTP attempt and how long to wait.

    Pure function, no side effects. Caller sleeps `delay` seconds before retry.
    Returns (should_retry, delay_seconds).
    """
    if attempt >= max_retries:
        return False, 0.0

    backoff = float(min(2 ** attempt, 60))

    if exception is not None:
        return True, backoff

    if status_code == 429:
        if retry_after is not None:
            try:
                return True, min(max(0.0, float(retry_after)), 60.0)
            except (TypeError, ValueError):
                pass
        return True, backoff

    if status_code is not None and status_code >= 500:
        return True, backoff

    return False, 0.0
