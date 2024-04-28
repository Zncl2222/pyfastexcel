class CreateFileNotCalledError(Exception):
    """
    Exception raised when create_file function is not called
    before saving a file.
    """

    default_message = (
        'Function read_lib_and_create_excel should be' + 'called before saving the file.'
    )

    def __init__(self, message: str = default_message):
        self.message = message
        super().__init__(self.message)
