import os
import sys
# Dynamically add the root directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from step1_pyxis_processor import main as run_pyxis_processor

def main():
    print("=" * 60)
    print("CONTROLLED SUBSTANCE INFUSION RECONCILIATION PIPELINE")
    print("=" * 60)
    
    try:
        run_pyxis_processor()
    except Exception as e:
        print(f"\nERROR: Reconciliation pipeline failed!")
        print(f"Reason: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
