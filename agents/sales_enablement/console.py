"""Console commands for Sales Enablement agent.

4 independent steps:
  /scan     - Extract client names from Excel (pure Python, instant)
  /research - Enrich client profiles via LLM + web search
  /analyze  - Analyze PE firm portfolio via LLM + web search
  /report   - Generate PPTX from saved analysis (pure Python, instant)
"""

import shlex
from typing import cast

from rich.console import Console

from agent_kit.api.console.server import SlashCommands

from .agent import SalesEnablementAgent
import json as _json

from .tools import (
    scan_client_list, generate_report_pptx, generate_portco_report_pptx,
    get_flagged_profiles, update_client_profile, replace_client_profile,
    load_clients_for_research, load_client_profiles_for_comps,
    get_pending_portfolio_companies, CLIENT_LIST_FILE,
)


class SalesEnablementCommands(SlashCommands):
    """Console commands for the Sales Enablement agent."""

    def __init__(self, console: Console):
        super().__init__(console)

        self.register_command("/scan", self._handle_scan, "Extract client names from Excel",
                              "Extract client names from input/client list.xlsx (Step 1)\nUsage: /scan")

        self.register_command("/research", self._handle_research, "Research and classify clients",
                              "Enrich client profiles with industry/sector data (Step 2)\nUsage: /research")

        self.register_command("/analyze-pe", self._handle_analyze_pe, "Analyze a PE firm's portfolio",
                              "Analyze full PE firm portfolio and save results\nUsage: /analyze-pe Gauge Capital")

        self.register_command("/analyze-portco", self._handle_analyze_portco, "Analyze a single company",
                              "Analyze a single portfolio company\nUsage: /analyze-portco Leaf Home")

        self.register_command("/review", self._handle_review, "Review flagged client profiles",
                              "Review low/medium confidence profiles interactively\nUsage: /review")

        self.register_command("/report-pe", self._handle_report_pe, "Generate PE-level PowerPoint",
                              "Generate PE-level PPTX from saved analysis\nUsage: /report-pe Gauge Capital")

        self.register_command("/report-portco", self._handle_report_portco, "Generate company-level PowerPoint",
                              "Generate company-level PPTX from saved analysis\nUsage: /report-portco Leaf Home")

        self.register_command("/list", self._handle_list, "Show status of all data",
                              "Show client profiles status and available PE firm analyses")

    def _parse_args(self, user_input: str) -> tuple[str, list[str]]:
        """Parse input preserving quoted strings."""
        try:
            parts = shlex.split(user_input)
        except ValueError:
            parts = user_input.strip().split()
        cmd = parts[0].lower() if parts else ""
        return cmd, parts[1:]

    async def handle_input(self, user_input: str) -> bool:
        """Route commands with proper arg parsing."""
        if not user_input.startswith("/"):
            await self._handle_chat(user_input)
            return True

        cmd, args = self._parse_args(user_input)

        handlers = {
            "/scan": self._handle_scan, "/research": self._handle_research,
            "/review": self._handle_review,
            "/analyze-pe": self._handle_analyze_pe, "/analyze-portco": self._handle_analyze_portco,
            "/report-pe": self._handle_report_pe, "/report-portco": self._handle_report_portco,
            "/list": self._handle_list,
        }

        if cmd in handlers:
            await handlers[cmd](args)
            return True

        if await super().handle_input(user_input):
            return True

        self.console.print(f"[red]Unknown command: {cmd}[/red]")
        self.console.print("Type [cyan]/help[/cyan] for available commands")
        return True

    # ── Step 1: Scan ──────────────────────────────────────────────

    async def _handle_scan(self, args: list[str]) -> None:
        """Extract client names from Excel. No LLM needed."""
        self.console.print("\n[bold]Scanning client list...[/bold]")
        self.console.print(f"[dim]  Source: {CLIENT_LIST_FILE}[/dim]\n")

        result = scan_client_list()

        if "error" in result:
            self.console.print(f"[red]Error: {result['error']}[/red]")
            return

        self.console.print("[green]Scan complete:[/green]")
        self.console.print(f"  Total clients: {result['total']}")
        self.console.print(f"  New:           {result['new']}")
        self.console.print(f"  Existing:      {result['existing']}")
        self.console.print(f"  Saved to:      {result['file']}\n")

    # ── Step 2: Research ──────────────────────────────────────────

    async def _handle_research(self, args: list[str]) -> None:
        """Research clients via LLM -- one client per LLM call, looped in Python."""
        if not self.session_id:
            self.console.print("[red]Session not initialized[/red]")
            return

        client_data = await load_clients_for_research()
        if "error" in client_data:
            self.console.print(f"[red]Error: {client_data['error']}[/red]")
            return

        pending = client_data["pending_clients"]
        already = client_data["already_profiled"]
        total = client_data["total_clients"]

        if not pending:
            self.console.print(f"[green]All {total} clients are already profiled. Nothing to do.[/green]\n")
            return

        self.console.print(f"\n[bold]Starting client research: {len(pending)} to profile ({already} already done)[/bold]\n")

        try:
            session = await self.session_store.get_session(self.session_id)
            if not session:
                self.console.print("[red]Session not found[/red]")
                return

            agent = cast(SalesEnablementAgent, await session.use_agent(SalesEnablementAgent))

            success = 0
            errors = 0
            for i, client in enumerate(pending, 1):
                name = client["name"]
                pe = client.get("pe_sponsor", "N/A")
                self.console.print(f"[dim]  [{already + i}/{total}] Researching: {name} (PE: {pe})...[/dim]")

                try:
                    await agent.research_client(name, pe)
                    success += 1
                except Exception as e:
                    errors += 1
                    self.console.print(f"[red]    Error on {name}: {e}[/red]")

            self.console.print(f"\n[bold green]Research Complete:[/bold green]")
            self.console.print(f"  Profiled: {success}")
            if errors:
                self.console.print(f"  Errors:   {errors}")
            self.console.print(f"  Total:    {already + success}/{total}")
            self.console.print(f"\nRun [cyan]/review[/cyan] to check flagged profiles.\n")
        except Exception as e:
            self.console.print(f"[red]Error: {e}[/red]")

    # ── Review: Fix flagged profiles ────────────────────────────────

    async def _handle_review(self, args: list[str]) -> None:
        """Interactive review of low/medium confidence profiles."""
        import asyncio
        import json

        flagged = get_flagged_profiles()
        if not flagged:
            self.console.print("[green]No flagged profiles to review. All profiles are high confidence.[/green]\n")
            return

        self.console.print(f"\n[bold]{len(flagged)} profiles flagged for review:[/bold]\n")

        for i, p in enumerate(flagged, 1):
            self.console.print(f"[bold yellow]--- {i}/{len(flagged)}: {p['name']} ---[/bold yellow]")
            self.console.print(f"  PE Sponsor:  {p.get('pe_sponsor', 'N/A')}")
            self.console.print(f"  Industry:    {p.get('industry', 'N/A')}")
            self.console.print(f"  Sector:      {p.get('sector', 'N/A')}")
            self.console.print(f"  Description: {p.get('description', 'N/A')}")
            self.console.print(f"  Revenue:     {p.get('estimated_revenue_range', 'N/A')}")
            self.console.print(f"  Confidence:  {p.get('confidence', 'N/A')}")
            self.console.print(f"  Notes:       {p.get('notes', '')}")
            self.console.print()
            self.console.print("[dim]  Options: Enter corrected company name, 'skip' to keep as-is, or 'ok' to mark high confidence[/dim]")

            # Read user input directly
            try:
                loop = asyncio.get_event_loop()
                user_input = await loop.run_in_executor(None, lambda: input("  > ").strip())
            except (EOFError, KeyboardInterrupt):
                self.console.print("\n[dim]Review cancelled.[/dim]\n")
                return

            if not user_input or user_input.lower() == "skip":
                self.console.print("[dim]  Skipped.[/dim]\n")
                continue

            if user_input.lower() == "ok":
                result = await update_client_profile(p["name"], json.dumps({"confidence": "high"}))
                if "error" in result:
                    self.console.print(f"[red]  Error: {result['error']}[/red]\n")
                else:
                    self.console.print(f"[green]  Marked as high confidence.[/green]\n")
                continue

            # User provided a corrected company name -- re-research via LLM then replace old profile
            corrected_name = user_input
            old_name = p["name"]
            pe_sponsor = p.get("pe_sponsor", "unknown")
            self.console.print(f"[dim]  Replacing '{old_name}' with '{corrected_name}'. Re-researching...[/dim]")

            if not self.session_id:
                self.console.print("[red]  Session not initialized.[/red]")
                continue

            try:
                session = await self.session_store.get_session(self.session_id)
                if not session:
                    self.console.print("[red]  Session not found.[/red]")
                    continue

                agent = cast(SalesEnablementAgent, await session.use_agent(SalesEnablementAgent))
                response = await agent.research(
                    f"Research this single client: '{corrected_name}' (PE sponsor: {pe_sponsor}). "
                    f"Web search and call save_single_client_profile with the result. One client only."
                )

                # Now replace the old profile with the newly saved one
                from .tools import PROFILES_JSON
                if PROFILES_JSON.exists():
                    with open(PROFILES_JSON) as f:
                        all_profiles = json.load(f)
                    # Find the new profile that was just saved
                    new_profile = next((pr for pr in all_profiles if pr["name"].lower() == corrected_name.lower()), None)
                    if new_profile:
                        new_profile["pe_sponsor"] = pe_sponsor
                        new_profile["confidence"] = "high"
                        new_profile["notes"] = f"Corrected from '{old_name}'"
                        result = replace_client_profile(old_name, new_profile)
                        self.console.print(f"[green]  Replaced '{old_name}' -> '{corrected_name}'[/green]\n")
                    else:
                        self.console.print(f"[yellow]  Researched but could not find saved profile for '{corrected_name}'. Old profile unchanged.[/yellow]\n")
                else:
                    self.console.print(f"[red]  Profiles file missing.[/red]\n")
            except Exception as e:
                self.console.print(f"[red]  Error: {e}[/red]\n")

        self.console.print("[bold green]Review complete.[/bold green]\n")

    # ── Step 3: Analyze ───────────────────────────────────────────

    async def _handle_analyze_pe(self, args: list[str]) -> None:
        """Analyze PE firm portfolio -- Python-driven loop, one company at a time."""
        if not args:
            self.console.print("[dim]Usage: /analyze-pe <PE firm name>\nExample: /analyze-pe Gauge Capital[/dim]")
            return

        if not self.session_id:
            self.console.print("[red]Session not initialized[/red]")
            return

        pe_firm = " ".join(args)

        try:
            session = await self.session_store.get_session(self.session_id)
            if not session:
                self.console.print("[red]Session not found[/red]")
                return

            agent = cast(SalesEnablementAgent, await session.use_agent(SalesEnablementAgent))

            # Check for resume -- do we already have portfolio companies?
            all_companies, done_names = get_pending_portfolio_companies(pe_firm)

            if not all_companies:
                # Step 1: Find portfolio companies
                self.console.print(f"\n[bold]Step 1: Finding active portfolio companies for {pe_firm}[/bold]")
                self.console.print("[dim]  (searching the web...)[/dim]\n")
                await agent.find_portfolio(pe_firm)

                all_companies, done_names = get_pending_portfolio_companies(pe_firm)
                if not all_companies:
                    self.console.print("[red]No portfolio companies found. Check the PE firm name.[/red]")
                    return

                self.console.print(f"[green]  Found {len(all_companies)} portfolio companies[/green]\n")
            else:
                self.console.print(f"\n[bold]Resuming analysis for {pe_firm}[/bold]")
                self.console.print(f"[dim]  {len(all_companies)} companies, {len(done_names)} already analyzed[/dim]\n")

            # Step 2: Analyze each company
            clients_summary = load_client_profiles_for_comps()
            pending = [c for c in all_companies if c.get("name", "").lower() not in done_names]

            if pending:
                self.console.print(f"[bold]Step 2: Analyzing {len(pending)} companies (comps + savings)[/bold]\n")

                for i, company in enumerate(pending, len(done_names) + 1):
                    name = company.get("name", "Unknown")
                    self.console.print(f"  [dim][{i}/{len(all_companies)}] {name}...[/dim]")
                    try:
                        await agent.analyze_company(pe_firm, company, clients_summary)
                        self.console.print(f"  [green]  [{i}/{len(all_companies)}] {name}: done[/green]")
                    except Exception as e:
                        self.console.print(f"  [red]  Error on {name}: {e}[/red]")
            else:
                self.console.print("[dim]  All companies already analyzed.[/dim]")

            # Step 3: Generate summary
            self.console.print(f"\n[bold]Step 3: Generating executive summary[/bold]\n")
            await agent.generate_summary(pe_firm)

            self.console.print("[bold green]Analysis Complete![/bold green]")
            self.console.print(f"Run [cyan]/report-pe {pe_firm}[/cyan] to generate the PowerPoint.\n")

        except Exception as e:
            self.console.print(f"[red]Error: {e}[/red]")
            self.console.print("[dim]Progress has been saved. Run /analyze-pe again to resume.[/dim]\n")

    # ── Step 4: Report ────────────────────────────────────────────

    async def _handle_report_pe(self, args: list[str]) -> None:
        """Generate PE-level PPTX from saved analysis. No LLM needed."""
        if not args:
            self.console.print("[dim]Usage: /report-pe <PE firm name>\nExample: /report-pe Gauge Capital[/dim]")
            return

        pe_firm = " ".join(args)
        self.console.print(f"\n[bold]Generating report for: {pe_firm}[/bold]\n")

        result = generate_report_pptx(pe_firm)

        if "error" in result:
            self.console.print(f"[red]Error: {result['error']}[/red]")
            return

        self.console.print("[green]Report generated:[/green]")
        self.console.print(f"  Slides: {result['slide_count']}")
        self.console.print(f"  Output: {result['output_path']}\n")

    # ── Portco: Analyze single company ──────────────────────────────

    async def _handle_analyze_portco(self, args: list[str]) -> None:
        """Analyze a single portfolio company."""
        if not args:
            self.console.print("[dim]Usage: /analyze-portco <company name>\nExample: /analyze-portco Leaf Home[/dim]")
            return

        if not self.session_id:
            self.console.print("[red]Session not initialized[/red]")
            return

        company_name = " ".join(args)
        self.console.print(f"\n[bold]Analyzing company: {company_name}[/bold]")
        self.console.print("[dim]  (web searching for company details, comps, and savings...)[/dim]\n")

        try:
            session = await self.session_store.get_session(self.session_id)
            if not session:
                self.console.print("[red]Session not found[/red]")
                return

            agent = cast(SalesEnablementAgent, await session.use_agent(SalesEnablementAgent))
            clients_summary = load_client_profiles_for_comps()
            await agent.analyze_portco(company_name, clients_summary)

            # Ensure the file is saved under the user's original name
            from .tools import _pe_slug, DATA_DIR
            slug = _pe_slug(company_name)
            expected = DATA_DIR / f"portco_{slug}.json"
            if not expected.exists():
                # LLM may have saved under a different name -- find and rename
                import json as j2
                for f in DATA_DIR.glob("portco_*.json"):
                    if f == expected:
                        continue
                    try:
                        with open(f) as fh:
                            d = j2.load(fh)
                        if d.get("company_name", "").lower().replace(" ", "") in company_name.lower().replace(" ", "") or company_name.lower().replace(" ", "") in d.get("company_name", "").lower().replace(" ", ""):
                            f.rename(expected)
                            break
                    except Exception:
                        pass

            self.console.print("[bold green]Analysis Complete![/bold green]")
            self.console.print(f"Run [cyan]/report-portco {company_name}[/cyan] to generate the PowerPoint.\n")
        except Exception as e:
            self.console.print(f"[red]Error: {e}[/red]\n")

    # ── Portco: Generate report ───────────────────────────────────

    async def _handle_report_portco(self, args: list[str]) -> None:
        """Generate portco-level PPTX. No LLM needed."""
        if not args:
            self.console.print("[dim]Usage: /report-portco <company name>\nExample: /report-portco Leaf Home[/dim]")
            return

        company_name = " ".join(args)
        self.console.print(f"\n[bold]Generating portco report for: {company_name}[/bold]\n")

        result = generate_portco_report_pptx(company_name)

        if "error" in result:
            self.console.print(f"[red]Error: {result['error']}[/red]")
            return

        self.console.print("[green]Report generated:[/green]")
        self.console.print(f"  Slides: {result['slide_count']}")
        self.console.print(f"  Output: {result['output_path']}\n")

    # ── List status ────────────────────────────────────────────────

    async def _handle_list(self, args: list[str]) -> None:
        """Show status of all data files."""
        import json as j
        from .tools import CLIENTS_JSON, PROFILES_JSON, DATA_DIR, OUTPUT_DIR

        self.console.print()

        # Client list
        if CLIENTS_JSON.exists():
            with open(CLIENTS_JSON) as f:
                data = j.load(f)
            self.console.print(f"[bold]Client List:[/bold] {data.get('count', 0)} clients (last scanned: {data.get('last_scanned', 'N/A')[:10]})")
        else:
            self.console.print("[bold]Client List:[/bold] [dim]Not scanned yet. Run /scan[/dim]")

        # Profiles
        if PROFILES_JSON.exists():
            with open(PROFILES_JSON) as f:
                profiles = j.load(f)
            high = sum(1 for p in profiles if p.get("confidence") == "high")
            med = sum(1 for p in profiles if p.get("confidence") == "medium")
            low = sum(1 for p in profiles if p.get("confidence") == "low")
            self.console.print(f"[bold]Client Profiles:[/bold] {len(profiles)} profiled ({high} high, {med} medium, {low} low)")
        else:
            self.console.print("[bold]Client Profiles:[/bold] [dim]Not researched yet. Run /research[/dim]")

        # Analyses
        analysis_files = sorted(DATA_DIR.glob("analysis_*.json"))
        if analysis_files:
            self.console.print(f"\n[bold]Analyzed PE Firms:[/bold]")
            for af in analysis_files:
                with open(af) as f:
                    analysis = j.load(f)
                pe_name = analysis.get("pe_firm", {}).get("name", af.stem)
                n_companies = len(analysis.get("portfolio_companies", []))
                total_savings = sum(a.get("total_estimated_savings_mm", 0) for a in analysis.get("savings_assessment", []))
                report_path = OUTPUT_DIR / f"{af.stem.replace('analysis_', '')}_report.pptx"
                has_report = "[green]report generated[/green]" if report_path.exists() else "[dim]no report yet[/dim]"
                self.console.print(f"  {pe_name}: {n_companies} companies, ${total_savings:,.1f}MM savings -- {has_report}")
        else:
            self.console.print(f"\n[bold]Analyzed PE Firms:[/bold] [dim]None yet. Run /analyze-pe <PE firm>[/dim]")

        # Portco analyses
        portco_files = sorted(DATA_DIR.glob("portco_*.json"))
        if portco_files:
            self.console.print(f"\n[bold]Analyzed Companies:[/bold]")
            for pf in portco_files:
                with open(pf) as f:
                    pa = j.load(f)
                co_name = pa.get("company_name", pf.stem)
                rev = pa.get("estimated_revenue_mm", 0)
                sav = pa.get("total_estimated_savings_mm", 0)
                slug = pf.stem.replace("portco_", "")
                report_path = OUTPUT_DIR / f"portco_{slug}_report.pptx"
                has_report = "[green]report generated[/green]" if report_path.exists() else "[dim]no report yet[/dim]"
                self.console.print(f"  {co_name}: ${rev:,.0f}MM rev, ${sav:,.1f}MM savings -- {has_report}")

        self.console.print()

    # ── Chat fallback ─────────────────────────────────────────────

    async def _handle_chat(self, user_input: str) -> None:
        """Handle free-form chat. Continues the last active step's conversation."""
        self.console.print("[dim]Chat is not supported between steps. Use a command:[/dim]")
        self.console.print("[dim]  /scan, /research, /analyze-pe, /analyze-portco, /report-pe, /report-portco[/dim]")
