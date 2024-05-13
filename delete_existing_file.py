import os


def remove_files_from_folder(file_names, folder_path, logger):
    # Delete each leave register file
    for file in file_names:
        file_path = os.path.join(folder_path, file)
        os.remove(file_path)
        logger.info(f"Deleted: {file_path}")


def delete_files(logger, folder_path):
    for folder in folder_path:
        # List all files in the specified folder
        all_files = os.listdir(folder)
        # Filter files that contain "Leave Register" in their name
        leave_register_files = [file for file in all_files if "Leave Register" in file]
        remove_files_from_folder(leave_register_files, folder, logger)
        attendance_files = [file for file in all_files if "Attendance" in file]
        remove_files_from_folder(attendance_files, folder, logger)
