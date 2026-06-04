#!/usr/bin/env python3
"""Trigger motherboard front-panel functions through two GPIO-driven optocouplers.

Actions:
- reset-cmos    -> 15 second pulse on CMOS pin
- soft-power    -> 1 second pulse on power pin
- hard-poweroff -> 7 second pulse on power pin
- boot-list     -> print GPIO pin list and boot states
- boot-apply    -> write GPIO boot config lines to config.txt

Example:
	python3 KVM_GPIO.py soft-power
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time

try:
	import RPi.GPIO as GPIO
except ImportError:  # pragma: no cover - only hits when not on a Pi
	GPIO = None


# BCM pin numbers (GPIOxx). Change these to your actual wiring.
CMOS_RESET_PIN = 23
POWER_PIN = 24

# Pulse durations in seconds.
CMOS_RESET_SECONDS = 15.0
SOFT_POWER_SECONDS = 1.0
HARD_POWEROFF_SECONDS = 7.0

# Many optocoupler input boards are active-low (IN pulled low = on).
# If your channel LED turns on when GPIO is LOW, keep this False.
ACTIVE_HIGH = False

# 40-pin header BCM GPIO numbers (common user-available pins).
HEADER_GPIO_PINS = (
	2,
	3,
	4,
	5,
	6,
	7,
	8,
	9,
	10,
	11,
	12,
	13,
	14,
	15,
	16,
	17,
	18,
	19,
	20,
	21,
	22,
	23,
	24,
	25,
	26,
	27,
)

DEFAULT_BOOT_CONFIG_PATH = "/boot/firmware/config.txt"


def require_gpio() -> None:
	if GPIO is None:
		raise SystemExit(
			"RPi.GPIO is not available. Install it on Raspberry Pi OS with: "
			"sudo apt install python3-rpi.gpio"
		)


def active_level() -> int:
	return GPIO.HIGH if ACTIVE_HIGH else GPIO.LOW


def inactive_level() -> int:
	return GPIO.LOW if ACTIVE_HIGH else GPIO.HIGH


def inactive_pull() -> int:
	return GPIO.PUD_DOWN if ACTIVE_HIGH else GPIO.PUD_UP


def setup_pins() -> None:
	require_gpio()
	assert GPIO is not None
	GPIO.setwarnings(False)
	GPIO.setmode(GPIO.BCM)
	# Bias each pin to inactive first, then switch to output at inactive level.
	for pin in (CMOS_RESET_PIN, POWER_PIN):
		GPIO.setup(pin, GPIO.IN, pull_up_down=inactive_pull())
		GPIO.setup(pin, GPIO.OUT, initial=inactive_level())


def pulse(pin: int, duration_seconds: float, label: str) -> None:
	require_gpio()
	assert GPIO is not None
	print(f"[{label}] activating GPIO {pin} for {duration_seconds:.1f} seconds...")
	GPIO.output(pin, active_level())
	time.sleep(duration_seconds)
	GPIO.output(pin, inactive_level())
	print(f"[{label}] complete.")


def parse_pin_list(pin_text: str) -> list[int]:
	pins: list[int] = []
	for part in pin_text.split(","):
		chunk = part.strip()
		if not chunk:
			continue
		if not chunk.isdigit():
			raise ValueError(f"Invalid pin '{chunk}'. Use comma-separated BCM integers.")
		pin = int(chunk)
		if pin < 0 or pin > 27:
			raise ValueError(f"GPIO {pin} is out of range. Expected BCM 0-27.")
		pins.append(pin)

	if not pins:
		raise ValueError("No GPIO pins were provided.")

	return sorted(set(pins))


def resolve_boot_pins(args: argparse.Namespace) -> list[int]:
	if args.all_pins or args.pins.strip().lower() == "all":
		return list(HEADER_GPIO_PINS)
	return parse_pin_list(args.pins)


def boot_line(pin: int, state: str) -> str:
	# dh/dl sets the boot-time idle (rest) level for the pin.
	return f"gpio={pin}=op,{state}"


def parse_existing_boot_states(lines: list[str]) -> dict[int, str]:
	pattern = re.compile(r"^\s*gpio\s*=\s*(\d+)\s*=\s*([^#\r\n]+)", re.IGNORECASE)
	states: dict[int, str] = {}

	for line in lines:
		match = pattern.match(line)
		if not match:
			continue

		pin = int(match.group(1))
		tokens = [token.strip().lower() for token in match.group(2).split(",")]
		state = "set"
		for token in tokens:
			if token in ("dh", "dl"):
				state = token
		states[pin] = state

	return states


def list_boot_template(config_path: str, state: str, show_template: bool) -> int:
	print("BCM GPIO pins on Raspberry Pi 40-pin header:")
	print(", ".join(str(pin) for pin in HEADER_GPIO_PINS))
	print(
		"Note: dh/dl sets boot idle state. Trigger actions in this script pulse to "
		"the opposite level, then return to idle."
	)

	if os.path.exists(config_path):
		with open(config_path, "r", encoding="utf-8") as handle:
			existing_states = parse_existing_boot_states(handle.readlines())
		print(f"\nCurrent boot states from {config_path}:")
		for pin in HEADER_GPIO_PINS:
			current = existing_states.get(pin, "(not set)")
			print(f"GPIO {pin:>2}: {current}")
	else:
		print(f"\nConfig file not found at {config_path}; no current states to display.")

	if show_template:
		print(f"\nTemplate lines for state '{state}':")
		for pin in HEADER_GPIO_PINS:
			print(boot_line(pin, state))
	else:
		print("\nTemplate lines hidden. Use --show-template to display them.")

	return 0


def apply_boot_config(config_path: str, pins: list[int], state: str, dry_run: bool) -> int:
	desired = {pin: boot_line(pin, state) for pin in pins}
	pattern = re.compile(r"^\s*gpio\s*=\s*(\d+)\s*=.*$", re.IGNORECASE)

	if os.path.exists(config_path):
		with open(config_path, "r", encoding="utf-8") as handle:
			lines = handle.readlines()
	else:
		lines = []

	updated: list[str] = []
	replaced: set[int] = set()

	for line in lines:
		match = pattern.match(line)
		if match:
			pin = int(match.group(1))
			if pin in desired:
				if pin not in replaced:
					updated.append(desired[pin] + "\n")
					replaced.add(pin)
				continue
		updated.append(line)

	for pin in pins:
		if pin not in replaced:
			updated.append(desired[pin] + "\n")

	if dry_run:
		print("Dry run; no file changes written.")
		print(f"Would write these pin defaults to {config_path}:")
		for pin in pins:
			print(desired[pin])
		return 0

	dir_name = os.path.dirname(config_path)
	if dir_name:
		os.makedirs(dir_name, exist_ok=True)

	try:
		with open(config_path, "w", encoding="utf-8") as handle:
			handle.writelines(updated)
	except PermissionError:
		print(f"Permission denied writing {config_path}.")
		print("Re-run with sudo for boot config changes, for example:")
		print("sudo python3 KVM_GPIO.py boot-apply --pins 23,24 --state dh")
		return 1

	print(f"Updated {config_path} for {len(pins)} pin(s) with state '{state}'.")
	print("Reboot Raspberry Pi for boot-time defaults to take effect.")
	return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description=(
			"Control motherboard power/reset lines through optocouplers and "
			"manage Raspberry Pi boot-time GPIO defaults."
		)
	)
	parser.add_argument(
		"action",
		choices=(
			"reset-cmos",
			"soft-power",
			"hard-poweroff",
			"boot-list",
			"boot-apply",
		),
		help="Action to run.",
	)
	parser.add_argument(
		"--state",
		choices=("dh", "dl"),
		default="dh",
		help=(
			"Boot idle level: dh=drive high (about 3.3V), dl=drive low (about 0V). "
			"Trigger pulses toggle to the opposite level briefly."
		),
	)
	parser.add_argument(
		"--pins",
		default=f"{CMOS_RESET_PIN},{POWER_PIN}",
		help=(
			"Comma-separated BCM GPIO pins for boot-apply, or 'all' for full "
			"header list."
		),
	)
	parser.add_argument(
		"--all-pins",
		action="store_true",
		help="Use all common header GPIO pins (BCM 2-27, excluding 0/1).",
	)
	parser.add_argument(
		"--config-path",
		default=DEFAULT_BOOT_CONFIG_PATH,
		help="Path to Raspberry Pi config.txt file.",
	)
	parser.add_argument(
		"--dry-run",
		action="store_true",
		help="Preview boot-apply changes without writing to disk.",
	)
	parser.add_argument(
		"--show-template",
		action="store_true",
		help="With boot-list, also print template gpio= lines for the selected state.",
	)
	return parser.parse_args(argv)


def main(argv: list[str]) -> int:
	args = parse_args(argv)
	trigger_actions = {
		"reset-cmos": (CMOS_RESET_PIN, CMOS_RESET_SECONDS, "Reset CMOS"),
		"soft-power": (POWER_PIN, SOFT_POWER_SECONDS, "Soft power toggle"),
		"hard-poweroff": (POWER_PIN, HARD_POWEROFF_SECONDS, "Hard power off"),
	}

	if args.action in trigger_actions:
		pin, duration, label = trigger_actions[args.action]

		setup_pins()
		try:
			pulse(pin, duration, label)
		except KeyboardInterrupt:
			print("Interrupted; returning pins to inactive state.")
		finally:
			assert GPIO is not None
			GPIO.output(CMOS_RESET_PIN, inactive_level())
			GPIO.output(POWER_PIN, inactive_level())
			GPIO.cleanup()
		return 0

	if args.action == "boot-list":
		return list_boot_template(args.config_path, args.state, args.show_template)

	if args.action == "boot-apply":
		if hasattr(os, "geteuid") and os.geteuid() != 0 and not args.dry_run:
			print("boot-apply needs root privileges to write boot config files.")
			print("Re-run with sudo, or use --dry-run to preview changes.")
			return 1
		try:
			pins = resolve_boot_pins(args)
		except ValueError as exc:
			print(f"Pin parsing error: {exc}")
			return 2
		return apply_boot_config(args.config_path, pins, args.state, args.dry_run)

	return 0


if __name__ == "__main__":
	sys.exit(main(sys.argv[1:]))
