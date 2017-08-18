import os


def get_last_modified(directory):
    if os.path.isfile(directory):
        return os.path.getmtime(directory)
    else:
        return None


# does not check for circular dependencies
def _recurse_dependency_list(project_dir, dependency_dict, files_to_check):
    if files_to_check:
        for file in files_to_check:
            if "last_modified" not in dependency_dict[file] or dependency_dict[file]["last_modified"] is None:
                dependency_dict[file]["last_modified"] = get_last_modified(project_dir + file)
                # if dependency_dict[file]["last_modified"] is None:
                #     print("file: {}".format(file))

            dependency_dict = \
                _recurse_dependency_list(project_dir, dependency_dict, dependency_dict[file]["dependencies"])

    return dependency_dict


# resolve dependency change times
# if changes in dependencies, run change function
# return completed dictionary
# paths are in relation to the parent_dir
def change_check(project_dir, dependency_dict, files_to_check, change_func=None, required_items=None):
    dependency_dict = _recurse_dependency_list(project_dir, dependency_dict, files_to_check)

    for file in files_to_check:
        info = dependency_dict[file]
        last_modified = info["last_modified"]

        dep_changes = []

        for dependency in dependency_dict[file]["dependencies"]:
            dependency_info = dependency_dict[dependency]
            dep_last_modified = dependency_info["last_modified"]

            if last_modified is None or dep_last_modified is None or dep_last_modified - last_modified > 0.1:
                dep_changes.append(dependency)

        if dep_changes and change_func is not None:
            change_func(required_items, file, dep_changes)

        dependency_dict[file]["changes"] = dep_changes

    return dependency_dict

