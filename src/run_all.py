"""
run_all.py — Pipeline Orchestrator
=====================================
Runs the full e-commerce product data migration pipeline in order:

  1. getDataAuto.py          — Scrape product name, description, and specs
  2. imageScraper.py         — Download product images with automatic fallback
  3. measurementConverter.py — Add metric equivalents to all imperial measurements
  4. aiProcessor.py          — Generate SEO metadata via AI

Run individual steps by passing their number(s) as arguments:
  python run_all.py           → runs all 4 steps
  python run_all.py 1 3       → runs only step 1 and step 3
  python run_all.py 2         → runs only step 2
"""

import sys
import logging
import time

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def run_step(step_number: int, name: str, fn):
    log.info("=" * 60)
    log.info("STEP %d — %s", step_number, name)
    log.info("=" * 60)
    start = time.time()
    try:
        fn()
        elapsed = time.time() - start
        log.info("✓ Step %d done in %.1fs", step_number, elapsed)
    except Exception as e:
        log.error("✗ Step %d FAILED: %s", step_number, e)
        raise


def main():
    args = sys.argv[1:]
    requested = set(int(a) for a in args) if args else {1, 2, 3, 4}

    steps = {}

    if 1 in requested:
        from getDataAuto import main as scrape_data
        steps[1] = ("Scrape product data (name / description / specs)", scrape_data)

    if 2 in requested:
        from imageScraper import main as scrape_images
        steps[2] = ("Download product images", scrape_images)

    if 3 in requested:
        from measurementConverter import main as convert_measurements
        steps[3] = ("Convert measurements to metric", convert_measurements)

    if 4 in requested:
        from aiProcessor import main as run_ai
        steps[4] = ("AI processing (meta title / description / tags / brand)", run_ai)

    if not steps:
        log.error("No valid step numbers provided. Use 1–4.")
        sys.exit(1)

    log.info("Running steps: %s", sorted(steps.keys()))
    pipeline_start = time.time()

    for step_number in sorted(steps.keys()):
        name, fn = steps[step_number]
        run_step(step_number, name, fn)

    total = time.time() - pipeline_start
    log.info("=" * 60)
    log.info("Pipeline complete in %.1fs", total)
    log.info("=" * 60)


if __name__ == "__main__":
    main()