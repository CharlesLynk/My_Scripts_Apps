#!/usr/bin/env python3
"""Trigger motherboard front-panel functions through two GPIO-driven optocouplers.

Actions:
- reset-cmos    -> 15 second pulse on CMOS pin
- soft-power    -> 1 second pulse on power pin
- hard-poweroff -> 7 second pulse on power pin

Example:
	python3 KVM_GPIO.py soft-power
"""

from __future__ import annotations

import argparse
import sys
import time

try:
	import RPi.GPIO as GPIO
except ImportError as exc:  # pragma: no cover - only hits when not on a Pi
	raise SystemExit(
		"RPi.GPIO is not available. Install it on Raspberry Pi OS with: "
		"sudo apt install python3-rpi.gpio"
	) from exc


# BCM pin numbers (GPIOxx). Change these to your actual wiring.
CMOS_RESET_PIN = 20
POWER_PIN = 21

# Pulse durations in seconds.
CMOS_RESET_SECONDS = 15.0
SOFT_POWER_SECONDS = 1.0
HARD_POWEROFF_SECONDS = 7.0

# Most DIY optocoupler circuits are triggered by driving GPIO HIGH.
# Set ACTIVE_HIGH = False if your module triggers when GPIO goes LOW.
ACTIVE_HIGH = True


def active_level() -> int:
	return GPIO.HIGH if ACTIVE_HIGH else GPIO.LOW


def inactive_level() -> int:
	return GPIO.LOW if ACTIVE_HIGH else GPIO.HIGH


def setup_pins() -> None:
	GPIO.setwarnings(False)
	GPIO.setmode(GPIO.BCM)
	GPIO.setup(CMOS_RESET_PIN, GPIO.OUT, initial=inactive_level())
	GPIO.setup(POWER_PIN, GPIO.OUT, initial=inactive_level())


def pulse(pin: int, duration_seconds: float, label: str) -> None:
	print(f"[{label}] activating GPIO {pin} for {duration_seconds:.1f} seconds...")
	GPIO.output(pin, active_level())
	time.sleep(duration_seconds)
	GPIO.output(pin, inactive_level())
	print(f"[{label}] complete.")


def parse_args(argv: list[str]) -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description=(
			"Pulse Raspberry Pi GPIO pins to control motherboard front-panel "
			"power/reset lines through optocouplers."
		)
	)
	parser.add_argument(
		"action",
		choices=("reset-cmos", "soft-power", "hard-poweroff"),
		help="Action to trigger.",
	)
	return parser.parse_args(argv)


def main(argv: list[str]) -> int:
	args = parse_args(argv)

	actions = {
		"reset-cmos": (CMOS_RESET_PIN, CMOS_RESET_SECONDS, "Reset CMOS"),
		"soft-power": (POWER_PIN, SOFT_POWER_SECONDS, "Soft power toggle"),
		"hard-poweroff": (POWER_PIN, HARD_POWEROFF_SECONDS, "Hard power off"),
	}

	pin, duration, label = actions[args.action]

	setup_pins()
	try:
		pulse(pin, duration, label)
	except KeyboardInterrupt:
		print("Interrupted; returning pins to inactive state.")
	finally:
		GPIO.output(CMOS_RESET_PIN, inactive_level())
		GPIO.output(POWER_PIN, inactive_level())
		GPIO.cleanup()

	return 0


if __name__ == "__main__":
	sys.exit(main(sys.argv[1:]))
