import os
import subprocess
import questionary
from yaspin import yaspin
from yaspin.spinners import Spinners
import sys

# === Correct base directory ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# === Map of report scripts ===
SCRIPTS = {
    "Phase 1 - Detailed NVT Report": "phase_1/detailed_nvt_reports/detailed_nvt_reports.py",
    "Phase 1 - Executive Report": "phase_1/executive_report/executive_report.py",
    "Phase 1 - Front Page": "phase_1/front/front.py",
    "Phase 2 - Executive Report": "phase_2/executive_report/executive_report.py",
}

# === Scripts that use input() and need clean stdin/stdout ===
INTERACTIVE_SCRIPTS = {
    "Phase 1 - Front Page"
}

def run_script(script_rel_path, env, interactive=False):
    script_path = os.path.join(BASE_DIR, script_rel_path)
    if not os.path.exists(script_path):
        print(f"❌ Script not found: {script_path}")
        return

    if interactive:
        # Run normally without spinner for input()-based scripts
        print(f"\n▶️  Running: {script_path}\n")
        try:
            subprocess.run([sys.executable, script_path], check=True, env=env)
            print("✅ Done.\n")
        except subprocess.CalledProcessError:
            print("💥 Script execution failed.\n")
    else:
        # Use spinner for non-interactive scripts
        with yaspin(Spinners.earth, text="Running script...") as spinner:
            try:
                subprocess.run([sys.executable, script_path], check=True, env=env)
                spinner.ok("✅")
            except subprocess.CalledProcessError:
                spinner.fail("💥 Script execution failed")

def main():
    # Ask user for custom output folder name (once)
    output_folder = questionary.text("🗂️ Enter custom folder name for output files:").ask()
    if not output_folder:
        print("❌ No folder name provided. Exiting.")
        return

    # Prepare environment
    env = os.environ.copy()
    env["CUSTOM_OUTPUT_DIR"] = os.path.join(BASE_DIR, output_folder)

    while True:
        choice = questionary.select(
            "📊 Select a report script to run:",
            choices=list(SCRIPTS.keys()) + ["Exit"]
        ).ask()

        if choice == "Exit":
            print("👋 Exiting.")
            break

        run_script(SCRIPTS[choice], env, interactive=(choice in INTERACTIVE_SCRIPTS))

if __name__ == "__main__":
    main()
