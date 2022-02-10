import NXOpen

def nxprint(message):
    """
    Simple wrapper for writing to the listing window.
    Useful for debugging NX journals.
    @param  message     message to print to the listing window
    """
    nxSession = NXOpen.Session.GetSession()
    lw = nxSession.ListingWindow
    lw.Open()
    lw.WriteLine(str(message))
    lw.Close()

def main():
    pass

if __name__ == '__main__':
    main()
