def top_words(text, n):
    counts = {}
    for w in text.split(" "):
        if w in counts:
            counts[w] += 1
        else:
            counts[w] = 1
    return sorted(counts.items())
