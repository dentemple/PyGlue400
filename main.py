from as400 import AS400


IS_TESTING = True


def main():
    # Upload references for the main GUI
    # Build the main GUI
    pass


def test():
    # AS400 Test
    print("Starting the AS400 test...")
    print("Please ensure the AS400 is active and ready...")
    as400 = AS400()
    as400.set_connection(name="A")
    print(as400.is_ready())
    as400.set_cursor(row=1, col=1)
    print("The AS400 Test has completed successfully")


if __name__ == "__main__":
    if not IS_TESTING:
        main()
    else:
        test()
