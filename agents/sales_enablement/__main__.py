"""Entry point: python -m agents.sales_enablement"""

import asyncio
import sys

from agent_kit.config import setup_configuration
from agent_kit.utils import set_app_name
from agent_kit.api.console import run_console

from .console import SalesEnablementCommands


def main():
    try:
        set_app_name("sales-enablement")
        config = asyncio.run(setup_configuration())

        if config.interfaces.console.enabled:
            asyncio.run(run_console(SalesEnablementCommands))
        else:
            print("Error: No interfaces enabled.")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\nGoodbye!")
        sys.exit(0)


if __name__ == "__main__":
    main()
