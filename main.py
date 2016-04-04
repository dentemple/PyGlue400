from as400 import AS400


def main():
    print("Starting test...")
    as400 = AS400()
    as400.set_connection(name="A")
    print(as400.is_ready())
    as400.set_cursor(row=1, col=1)
    print("Test is completed")


if __name__ == "__main__":
    main()
