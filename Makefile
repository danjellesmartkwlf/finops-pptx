.PHONY: run report check-db sync clean help

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-15s\033[0m %s\n", $$1, $$2}'

run: ## Launch the Streamlit UI
	uv run streamlit run app.py

report: ## Generate PPTX report (no UI). Usage: make report [MONTH=January] [YEAR=2026]
	uv run python generate_report.py -m $(or $(MONTH),January) -y $(or $(YEAR),2026)

check-db: ## Test Redshift connectivity
	uv run python -c "from src.ingestion import get_redshift_connection; conn = get_redshift_connection(); cur = conn.cursor(); cur.execute('SELECT 1'); print('Redshift connection: OK'); conn.close()"

sync: ## Install/sync dependencies
	uv sync

clean: ## Remove output files and Python cache
	rm -rf output/*.pptx __pycache__ src/__pycache__ .streamlit

inspect-template: ## Show PowerPoint template structure
	uv run python -c "from src.pptx_gen import inspect_template; import json; print(json.dumps(inspect_template('pptx_template/template1.pptx'), indent=2, default=str))"
