from hbx.retry import retry_policy


def test_exception_retries_with_backoff():
    should, delay = retry_policy(ValueError("boom"), None, None, 0, 4)
    assert should is True
    assert delay == 1.0


def test_backoff_doubles_per_attempt():
    assert retry_policy(ValueError(), None, None, 1, 4)[1] == 2.0
    assert retry_policy(ValueError(), None, None, 3, 8)[1] == 8.0


def test_backoff_caps_at_60():
    assert retry_policy(ValueError(), None, None, 7, 10)[1] == 60.0


def test_max_retries_stops():
    should, delay = retry_policy(ValueError(), None, None, 4, 4)
    assert should is False
    assert delay == 0.0


def test_429_honors_retry_after():
    should, delay = retry_policy(None, 429, "30", 0, 4)
    assert should is True
    assert delay == 30.0


def test_429_bad_retry_after_falls_back_to_backoff():
    should, delay = retry_policy(None, 429, "soon", 1, 4)
    assert should is True
    assert delay == 2.0


def test_500_retries():
    assert retry_policy(None, 500, None, 0, 4)[0] is True


def test_404_does_not_retry():
    assert retry_policy(None, 404, None, 0, 4)[0] is False


def test_401_does_not_retry():
    assert retry_policy(None, 401, None, 0, 4)[0] is False


def test_200_does_not_retry():
    assert retry_policy(None, 200, None, 0, 4)[0] is False


def test_429_small_retry_after_beats_backoff():
    should, delay = retry_policy(None, 429, "1", 3, 10)
    assert should is True
    assert delay == 1.0


def test_429_negative_retry_after_clamped_to_zero():
    assert retry_policy(None, 429, "-5", 0, 4) == (True, 0.0)
