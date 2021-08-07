import glob


def execute(target_der, exclusion_dir_list):
    all_list = glob.glob(target_der + "**/**.xlsx", recursive=True)
    return [f.replace('\\', '/')
            for f in all_list if is_target_dir(f, exclusion_dir_list)]


def is_target_dir(file_path, exclusion_dir_list):
    return not any(e in file_path for e in exclusion_dir_list)
