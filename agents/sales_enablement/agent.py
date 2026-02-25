"""Sales Enablement Agent.

LLM-powered steps, each called individually from Python loops in console.py:
  - research_client():        Research a single client
  - find_portfolio():         Find PE firm's active portfolio companies
  - analyze_company():        Analyze a single portfolio company (comps + savings)
  - generate_summary():       Generate executive summary from completed analysis
"""

import asyncio
import logging
from typing import Any

from agent_kit.agents.base_agent import BaseAgent
from agent_kit.api.progress import ProgressHandler
from agent_kit.clients.openai_client import OpenAIClient
from agent_kit.config.config import get_config

from .tools import (
    execute_tool,
    get_researcher_tool_definitions,
    get_portfolio_finder_tool_definitions,
    get_company_analyzer_tool_definitions,
    get_summary_tool_definitions,
    get_portco_analyzer_tool_definitions,
    load_savings_benchmarks,
    load_revenue_benchmarks,
)

logger = logging.getLogger(__name__)

MAX_RETRIES = 3


class SalesEnablementAgent(BaseAgent):
    """PE portfolio analysis and strategic sourcing assessment."""

    def __init__(self, openai_client: OpenAIClient, progress_handler: ProgressHandler):
        super().__init__(openai_client, progress_handler)

    def _max_iterations(self) -> int:
        config = get_config()
        agent_config = config.agent_configs.get(self.agent_type, {})
        return agent_config.get("max_iterations", config.agents.max_iterations)

    def _extract_text(self, response: Any) -> str:
        if hasattr(response, "output_text") and response.output_text:
            return response.output_text
        if hasattr(response, "output") and response.output:
            text_items = [i for i in response.output if hasattr(i, "type") and i.type == "text"]
            if text_items and hasattr(text_items[0], "text"):
                return text_items[0].text
        return ""

    async def _call_with_retry(self, func: Any, *args: Any, **kwargs: Any) -> Any:
        """Retry an async call up to MAX_RETRIES times on transient errors."""
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                return await func(*args, **kwargs)
            except Exception as e:
                if attempt == MAX_RETRIES:
                    raise
                logger.warning(f"Attempt {attempt}/{MAX_RETRIES} failed: {e}. Retrying in 3s...")
                await self.progress_handler.emit(f"Transient error, retrying ({attempt}/{MAX_RETRIES})...", "retry")
                await asyncio.sleep(3)
        raise RuntimeError("Unreachable")

    async def research_client(self, name: str, pe_sponsor: str) -> str:
        """Research a single client."""
        self.last_response_id = None
        prompts = self.render_prompt("sales_enablement", "researcher")
        query = (
            f"Research this company: '{name}' (PE sponsor: {pe_sponsor}). "
            f"Web search it, then call save_single_client_profile with the result."
        )
        response = await self._call_with_retry(
            self.execute_tool_conversation,
            instructions=prompts["instructions"],
            initial_input=[{"role": "user", "content": query}],
            tools=get_researcher_tool_definitions(),
            tool_executor=execute_tool,
            max_iterations=5,
            previous_response_id=None,
            response_format=None,
        )
        return self._extract_text(response)

    async def find_portfolio(self, pe_firm: str) -> str:
        """Find a PE firm's active portfolio companies. Returns JSON list via save tool."""
        self.last_response_id = None
        prompts = self.render_prompt("sales_enablement", "portfolio_finder")
        query = (
            f"Find all ACTIVE portfolio companies for '{pe_firm}'. "
            f"Web search, then call save_portfolio_companies with the results. "
            f"IMPORTANT: Use exactly '{pe_firm}' as the pe_firm_name parameter."
        )
        response = await self._call_with_retry(
            self.execute_tool_conversation,
            instructions=prompts["instructions"],
            initial_input=[{"role": "user", "content": query}],
            tools=get_portfolio_finder_tool_definitions(),
            tool_executor=execute_tool,
            max_iterations=10,
            previous_response_id=None,
            response_format=None,
        )
        return self._extract_text(response)

    async def analyze_company(self, pe_firm: str, company: dict[str, Any], clients_summary: str) -> str:
        """Analyze a single portfolio company: revenue, comps, savings."""
        self.last_response_id = None
        savings_bm = load_savings_benchmarks()
        revenue_bm = load_revenue_benchmarks()
        prompts = self.render_prompt(
            "sales_enablement", "company_analyzer",
            savings_benchmarks=savings_bm,
            revenue_benchmarks=revenue_bm,
        )
        query = (
            f"Analyze this portfolio company of '{pe_firm}':\n"
            f"  Name: {company.get('name', '')}\n"
            f"  Industry: {company.get('industry', 'Unknown')}\n"
            f"  Description: {company.get('description', '')}\n\n"
            f"Treya client profiles for comps matching:\n{clients_summary}\n\n"
            f"Web search for employee count and revenue data, run comps, estimate savings, "
            f"then call save_company_analysis with the results."
        )
        response = await self._call_with_retry(
            self.execute_tool_conversation,
            instructions=prompts["instructions"],
            initial_input=[{"role": "user", "content": query}],
            tools=get_company_analyzer_tool_definitions(),
            tool_executor=execute_tool,
            max_iterations=10,
            previous_response_id=None,
            response_format=None,
        )
        return self._extract_text(response)

    async def analyze_portco(self, company_name: str, clients_summary: str) -> str:
        """Analyze a single portfolio company standalone (not part of a PE portfolio run)."""
        self.last_response_id = None
        savings_bm = load_savings_benchmarks()
        revenue_bm = load_revenue_benchmarks()
        prompts = self.render_prompt(
            "sales_enablement", "company_analyzer",
            savings_benchmarks=savings_bm,
            revenue_benchmarks=revenue_bm,
        )
        query = (
            f"Analyze this company: '{company_name}'.\n\n"
            f"Treya client profiles for comps matching:\n{clients_summary}\n\n"
            f"Web search for employee count and revenue data, run comps, estimate savings, "
            f"then call save_portco_analysis with the results. "
            f"IMPORTANT: Use exactly '{company_name}' as the company_name parameter when calling save_portco_analysis."
        )
        response = await self._call_with_retry(
            self.execute_tool_conversation,
            instructions=prompts["instructions"],
            initial_input=[{"role": "user", "content": query}],
            tools=get_portco_analyzer_tool_definitions(),
            tool_executor=execute_tool,
            max_iterations=10,
            previous_response_id=None,
            response_format=None,
        )
        return self._extract_text(response)

    async def generate_summary(self, pe_firm: str) -> str:
        """Generate executive summary and finalize the analysis."""
        self.last_response_id = None
        prompts = self.render_prompt("sales_enablement", "summary_generator")
        query = (
            f"Generate an executive summary for the '{pe_firm}' analysis. "
            f"Call load_working_analysis to get all the data, write the summary, "
            f"then call finalize_analysis to save the completed report."
        )
        response = await self._call_with_retry(
            self.execute_tool_conversation,
            instructions=prompts["instructions"],
            initial_input=[{"role": "user", "content": query}],
            tools=get_summary_tool_definitions(),
            tool_executor=execute_tool,
            max_iterations=5,
            previous_response_id=None,
            response_format=None,
        )
        return self._extract_text(response)
