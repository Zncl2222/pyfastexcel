from pyfastexcel.validators import validate_call


def test_function_not_in_validators():
    @validate_call
    def test_function():
        pass

    test_function()
