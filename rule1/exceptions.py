class FieldNotFoundException(Exception):
    def __init__(self, section, title):
        self.section = section
        self.expected_title = title

        self.message = f'Cannot find title: "{title}" in section: "{section.value}"'
