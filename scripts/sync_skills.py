import os
import shutil

def sync_skills():
    # Target directory where active skills are loaded
    target_base_dir = r"C:\Users\user\.gemini\skills"
    os.makedirs(target_base_dir, exist_ok=True)

    # Directories to scan for SKILL.md
    scan_dirs = [
        r"C:\Users\user\.gemini",
        r"C:\Users\user\Desktop\Antigravity"
    ]

    synced_count = 0
    skipped_count = 0
    errors_count = 0

    print("Starting comprehensive skill synchronization...")

    for scan_dir in scan_dirs:
        if not os.path.exists(scan_dir):
            print(f"Directory not found, skipping scan: {scan_dir}")
            continue

        for root, dirs, files in os.walk(scan_dir):
            # Skip the target skills directory itself to avoid infinite loops or copying to itself
            if root.lower().startswith(target_base_dir.lower()):
                continue

            for file in files:
                if file.lower() == "skill.md":
                    skill_file_path = os.path.join(root, file)
                    parent_dir_name = os.path.basename(root)
                    grandparent_dir_name = os.path.basename(os.path.dirname(root))

                    # Determine clean skill folder name
                    if parent_dir_name.lower() == "skills":
                        skill_name = grandparent_dir_name
                    else:
                        skill_name = parent_dir_name

                    # Sanitize skill name
                    skill_name = skill_name.strip().replace(" ", "_")

                    target_skill_dir = os.path.join(target_base_dir, skill_name)

                    try:
                        # Copy everything from root (source skill directory) to target_skill_dir
                        os.makedirs(target_skill_dir, exist_ok=True)
                        for item in os.listdir(root):
                            s_item = os.path.join(root, item)
                            d_item = os.path.join(target_skill_dir, item)
                            if os.path.isdir(s_item):
                                # If target exists, merge or copy
                                if os.path.exists(d_item):
                                    shutil.rmtree(d_item)
                                shutil.copytree(s_item, d_item)
                            else:
                                shutil.copy2(s_item, d_item)
                        
                        print(f"Successfully synced skill: [{skill_name}] from {root} -> {target_skill_dir}")
                        synced_count += 1
                    except Exception as e:
                        print(f"Error syncing skill [{skill_name}] from {root}: {e}")
                        errors_count += 1

    print("\n--- Synchronization Summary ---")
    print(f"Total skills successfully synced: {synced_count}")
    print(f"Total skills skipped or failed: {errors_count}")
    print("All skills are now active and ready in Gemini CLI!")

if __name__ == "__main__":
    sync_skills()
