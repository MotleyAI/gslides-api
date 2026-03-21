import asyncio
import logging
import random
import time
from functools import wraps
from typing import Callable, Optional, Sequence, Type, TypeVar, Union

logger = logging.getLogger(__name__)

T = TypeVar("T")


def retry(
    func: Optional[Callable[..., T]] = None,
    args: tuple = (),
    kwargs: dict = None,
    *,
    max_attempts: int = 3,
    initial_delay: float = 1.0,
    max_delay: float = 60.0,
    exponential_base: float = 2.0,
    jitter: bool = True,
    exceptions: Union[Type[Exception], Sequence[Type[Exception]]] = Exception,
) -> Union[Callable[..., T], T]:
    """
    Can be used both as a decorator or a function to retry operations with exponential backoff.

    Can be used in two ways:
    1. As a decorator:
        @retry(max_attempts=3)
        def my_func():
            ...

    2. As a function:
        retry(my_func, kwargs={'param': 'value'}, max_attempts=3)

    Args:
        func: Function to retry (optional when used as decorator)
        args: Tuple of positional arguments to pass to the function (default: empty tuple)
        kwargs: Dictionary of keyword arguments to pass to the function (default: None)
        max_attempts: Maximum number of retry attempts (default: 3)
        initial_delay: Initial delay between retries in seconds (default: 1.0)
        max_delay: Maximum delay between retries in seconds (default: 60.0)
        exponential_base: Base for exponential backoff (default: 2.0)
        jitter: Whether to add random jitter to delays (default: True)
        exceptions: Exception or tuple of exceptions to catch and retry (default: Exception)

    Returns:
        The return value of the function if successful, or a decorator function if used as decorator

    Raises:
        The last exception encountered if all retries fail
    """
    kwargs = kwargs or {}

    def decorator(f):
        if asyncio.iscoroutinefunction(f):

            @wraps(f)
            async def wrapped_async(*a, **kw):
                return await _async_retry(
                    f,
                    args=a,
                    kwargs=kw,
                    max_attempts=max_attempts,
                    initial_delay=initial_delay,
                    max_delay=max_delay,
                    exponential_base=exponential_base,
                    jitter=jitter,
                    exceptions=exceptions,
                )

            return wrapped_async

        @wraps(f)
        def wrapped(*a, **kw):
            attempt = 0
            while attempt < max_attempts:
                try:
                    return f(*a, **kw)
                except exceptions as e:
                    attempt += 1

                    if attempt == max_attempts:
                        logger.error(
                            "Failed to execute %s after %d attempts. Final error: %s",
                            f.__name__,
                            max_attempts,
                            str(e),
                        )
                        raise

                    delay = min(initial_delay * (exponential_base ** (attempt - 1)), max_delay)

                    if jitter:
                        delay *= 1 + random.random()

                    logger.warning(
                        "Attempt %d/%d for %s failed: %s. Retrying in %.2f seconds...",
                        attempt,
                        max_attempts,
                        f.__name__,
                        str(e),
                        delay,
                    )

                    time.sleep(delay)

        return wrapped

    if func is None:
        return decorator
    return decorator(func)(*args, **kwargs)


async def _async_retry(
    func: Callable[..., T],
    args: tuple = (),
    kwargs: dict = None,
    *,
    max_attempts: int = 3,
    initial_delay: float = 1.0,
    max_delay: float = 60.0,
    exponential_base: float = 2.0,
    jitter: bool = True,
    exceptions: Union[Type[Exception], Sequence[Type[Exception]]] = Exception,
) -> T:
    kwargs = kwargs or {}

    attempt = 0
    while attempt < max_attempts:
        try:
            return await func(*args, **kwargs)
        except exceptions as e:
            attempt += 1

            if attempt == max_attempts:
                logger.error(
                    "Failed to execute %s after %d attempts. Final error: %s",
                    func.__name__,
                    max_attempts,
                    str(e),
                )
                raise

            delay = min(initial_delay * (exponential_base ** (attempt - 1)), max_delay)

            if jitter:
                delay *= 1 + random.random()

            logger.warning(
                "Attempt %d/%d for %s failed: %s. Retrying in %.2f seconds...",
                attempt,
                max_attempts,
                func.__name__,
                str(e),
                delay,
            )

            await asyncio.sleep(delay)
