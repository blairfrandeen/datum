def nxprint(message):
    """
    Simple wrapper for writing to the listing window.
    Useful for debugging NX journals.
    @param  message     message to print to the listing window
    """
    try:
        # TODO: See if I can set this up so I don't have to keep
        # getting the session & opening/closing the listing window
        # for every message. If printing hundreds of lines, this can
        # take a while.
        import NXOpen

        nxSession = NXOpen.Session.GetSession()
        lw = nxSession.ListingWindow
        lw.Open()
        lw.WriteLine(str(message))
        lw.Close()
    except ModuleNotFoundError:
        # if working outside NX, print messages to console
        print(message)

def nxdir(object):
    """Wrapper for dir() to print on separate lines"""
    nxprint(f"DIR FOR {object.__str__}:")
    for item in dir(object):
        nxprint(item)


def main():
    pass


if __name__ == "__main__":
    main()
