.PHONY: run report check-db sync clean help

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-15s\033[0m %s\n", $$1, $$2}'

run: ## Generate PPTX report (alias for `make report`)
	@$(MAKE) report

PREV_MONTH := $(shell date -v-1m +%B)
PREV_YEAR  := $(shell date -v-1m +%Y)

report: ## Generate PPTX report (no UI). Usage: make report [MONTH=March] [YEAR=2026]
	@OUT=$$(uv run python -u generate_report.py -m $(or $(MONTH),$(PREV_MONTH)) -y $(or $(YEAR),$(PREV_YEAR)) | tee /dev/tty | grep "^Done! Saved to" | sed 's/Done! Saved to //'); \
	open "$$OUT"

check-db: ## Test Redshift connectivity
	uv run python -c "from src.ingestion import get_redshift_connection; conn = get_redshift_connection(); cur = conn.cursor(); cur.execute('SELECT 1'); print('Redshift connection: OK'); conn.close()"

sync: ## Install/sync dependencies
	uv sync

clean: ## Remove output files and Python cache
	rm -rf output/*.pptx __pycache__ src/__pycache__

inspect-template: ## Show PowerPoint template structure
	uv run python -c "from src.pptx_gen import inspect_template; import json; print(json.dumps(inspect_template('pptx_template/template1.pptx'), indent=2, default=str))"
