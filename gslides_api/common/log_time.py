import asyncio
import functools
import logging
import time
from typing import Any, Callable, TypeVar

F = TypeVar("F", bound=Callable[..., object])


def log_time(func: F) -> F:
    """
    Decorator to log entry and exit of a function together with the execution time.
    """

    @functools.wraps(func)
    async def async_wrapper(*args, **kwargs) -> Any:
        # Get logger from the module where the decorated function is defined
        logger = logging.getLogger(func.__module__)

        start_time = time.time()
        logger.info("Entering %s", func.__qualname__)
        try:
            result = await func(*args, **kwargs)
            return result
        finally:
            elapsed_time = time.time() - start_time
            logger.info("Exiting %s (took %.2f s)", func.__qualname__, elapsed_time)

    @functools.wraps(func)
    def sync_wrapper(*args, **kwargs) -> Any:
        # Get logger from the module where the decorated function is defined
        logger = logging.getLogger(func.__module__)

        start_time = time.time()
        logger.info("Entering %s", func.__qualname__)
        try:
            result = func(*args, **kwargs)
            return result
        finally:
            elapsed_time = time.time() - start_time
            logger.info("Exiting %s (took %.2f s)", func.__qualname__, elapsed_time)

    return async_wrapper if asyncio.iscoroutinefunction(func) else sync_wrapper
