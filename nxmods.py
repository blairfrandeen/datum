def nxprint(message):
    """
    Simple wrapper for writing to the listing window.
    Useful for debugging NX journals.
    @param  message     message to print to the listing window
    """
    try:
        import NXOpen
        nxSession = NXOpen.Session.GetSession()
        lw = nxSession.ListingWindow
        lw.Open()
        lw.WriteLine(str(message))
        lw.Close()
    except ModuleNotFoundError:
        # if working outside NX, print messages to console
        print(message)

def main():
    pass

if __name__ == '__main__':
    main()
