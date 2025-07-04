import re
from datetime import datetime

def normalize_message(message):
    message = re.sub(
        r"Running script on (\d+x\d+) resolution with ([\d%]+) scale", 
        r"Running script on <reso> resolution with <scale> scale", 
        message
    )
    message = re.sub(
        r"(C:\\\\Users\\\\)[^\\\\]+(\\\\[^ ]*)", 
        r"\1<user>\2", 
        message
    )
    message = re.sub(
        r"https?://[^\s]+", 
        "<url>", 
        message
    )
    message = re.sub(
        r"Dragging mouse from \(\d+, \d+\) to \(\d+, \d+\)", 
        "Dragging mouse", 
        message
    )
    message = re.sub(
        r"Running method 'draw' on activity 'Microsoft Paint' with args: \{'points': \[\[.*?\]\], 'mouse_speed': \d+\}", 
        r"Running method 'draw' on activity 'Microsoft Paint' with args: {'points': <points_array>, 'mouse_speed': <mouse_speed>}", 
        message
    )
    message = re.sub(
        r"platform [^\s]+ -- Python [\d\.]+, pytest-[\d\.]+, pluggy-[\d\.]+ -- .*python.exe", 
        r"platform <platform> -- Python <py_ver>, pytest-<pytest_ver>, pluggy-<pluggy_ver> -- <python_exe>", 
        message
    )
    message = re.sub(
        r"(rootdir: ).+",
        r"\1<path>",
        message
    )
    message = re.sub(
        r"(?:.*?)(Test[^:]+::test_[^\s]+)\s+PASSED\s+\[\s*(\d+)%\]", 
        r"<path>::\1 PASSED [\2%]", 
        message
    )
    message = re.sub(
        r"\d+\.\d{2}s", 
        "<duration>s", 
        message
    )

    return message

def parse_log(log_lines):
    entries = []
    start_time = None
    current_timestamp = None
    current_message_lines = []

    for line in log_lines:
        match = re.match(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) - (.+)", line)
        if match:
            if current_timestamp is not None:
                full_message = "\n".join(current_message_lines)
                time_delta = (current_timestamp - start_time).total_seconds()
                entries.append((time_delta, normalize_message(full_message)))

            timestamp_str, raw_msg = match.groups()
            current_timestamp = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
            if start_time is None:
                start_time = current_timestamp
            current_message_lines = [raw_msg]
        else:
            if current_message_lines is not None:
                current_message_lines.append(line.strip())

    if current_timestamp is not None and current_message_lines:
        full_message = "\n".join(current_message_lines)
        time_delta = (current_timestamp - start_time).total_seconds()
        entries.append((time_delta, normalize_message(full_message)))

    return entries

def count_similarity(messages1, messages2):
    length = min(len(messages1), len(messages2))
    
    if length == 0:
        return 100.0

    same_lines = 0
    for i in range(length):
        line1 = messages1[i] if i < len(messages1) else None
        line2 = messages2[i] if i < len(messages2) else None
        if line1 == line2:
            same_lines += 1

    return (same_lines / length) * 100

def compare_logs(log1_lines, log2_lines):
    log1 = parse_log(log1_lines)
    log2 = parse_log(log2_lines)

    min_len = min(len(log1), len(log2))
    actions1 = [msg for _, msg in log1[:min_len]]
    actions2 = [msg for _, msg in log2[:min_len]]

    message_similarity = count_similarity(actions1, actions2)

    time1 = log1[min_len - 1][0] if min_len > 0 else 0
    time2 = log2[min_len - 1][0] if min_len > 0 else 0
    time_diff = abs(time1 - time2)
    plus_minus = "+" if time1 < time2 else "-"

    return {
        "message_similarity": round(message_similarity, 2),
        "time1": time1,
        "time2": time2,
        "plus_minus": plus_minus,
        "diff": time_diff
    }

def debug_similarity(log1_lines, log2_lines):
    log1 = parse_log(log1_lines)
    log2 = parse_log(log2_lines)

    max_len = max(len(log1), len(log2))
    debug_output = []

    debug_output.append("Line-by-line comparison of logs:")

    for i in range(max_len):
        entry1 = log1[i] if i < len(log1) else ("<no entry>", "<no entry>")
        entry2 = log2[i] if i < len(log2) else ("<no entry>", "<no entry>")

        timestamp1, msg1 = entry1
        timestamp2, msg2 = entry2

        if msg1 != msg2:
            debug_output.append(f"Line {i + 1}:")
            debug_output.append(f"  log1: [{timestamp1}] {msg1}")
            debug_output.append(f"  log2: [{timestamp2}] {msg2}")

    time1 = log1[-1][0] if log1 else 0
    time2 = log2[-1][0] if log2 else 0
    debug_output.append(f"\nLast timestamp in log1: {time1} seconds")
    debug_output.append(f"Last timestamp in log2: {time2} seconds")
    debug_output.append(f"Time difference: {abs(time1 - time2)} seconds")

    return "\n".join(debug_output)

if __name__ == "__main__":
    file_pairs = [
        ("example\\1.log", "1920x1080\\100%\\1.log"),
        ("example\\2.log", "1920x1080\\100%\\2.log"),
        ("example\\3.log", "1920x1080\\100%\\3.log"),
        ("example\\4.log", "1920x1080\\100%\\4.log"),
        ("example\\5.log", "1920x1080\\100%\\5.log"),
        
        ("example\\1.log", "1920x1080\\125%\\1.log"),
        ("example\\2.log", "1920x1080\\125%\\2.log"),
        ("example\\3.log", "1920x1080\\125%\\3.log"),
        ("example\\4.log", "1920x1080\\125%\\4.log"),
        ("example\\5.log", "1920x1080\\125%\\5.log"),
        
        ("example\\1.log", "1920x1080\\150%\\1.log"),
        ("example\\2.log", "1920x1080\\150%\\2.log"),
        ("example\\3.log", "1920x1080\\150%\\3.log"),
        ("example\\4.log", "1920x1080\\150%\\4.log"),
        ("example\\5.log", "1920x1080\\150%\\5.log"),

        ("example\\1.log", "1366x768\\100%\\1.log"),
        ("example\\2.log", "1366x768\\100%\\2.log"),
        ("example\\3.log", "1366x768\\100%\\3.log"),
        ("example\\4.log", "1366x768\\100%\\4.log"),
        ("example\\5.log", "1366x768\\100%\\5.log"),
        
        ("example\\1.log", "1366x768\\125%\\1.log"),
        ("example\\2.log", "1366x768\\125%\\2.log"),
        ("example\\3.log", "1366x768\\125%\\3.log"),
        ("example\\4.log", "1366x768\\125%\\4.log"),
        ("example\\5.log", "1366x768\\125%\\5.log"),

        ("example\\1.log", "1280x720\\100%\\1.log"),
        ("example\\2.log", "1280x720\\100%\\2.log"),
        ("example\\3.log", "1280x720\\100%\\3.log"),
        ("example\\4.log", "1280x720\\100%\\4.log"),
        ("example\\5.log", "1280x720\\100%\\5.log"),
    ]

    with open("comparison-results-line-per-line.txt", "w") as result_file, open("debug-output-line-per-line.txt", "w") as debug_file:
        for i, (file1, file2) in enumerate(file_pairs, 1):
            with open(file1) as f1, open(file2) as f2:
                log1_lines = f1.readlines()
                log2_lines = f2.readlines()

            result = compare_logs(log1_lines, log2_lines)
            debug = debug_similarity(log1_lines, log2_lines)

            result_file.write(f"=== Comparison {i}: {file1} vs {file2} ===\n")
            result_file.write(f"Message similarity: {result['message_similarity']}%\n")
            result_file.write(f"Time1: {result['time1']}s, Time2: {result['time2']}s, Diff: {result['plus_minus']} {result['diff']}s\n\n")

            debug_file.write(f"=== Debug {i}: {file1} vs {file2} ===\n")
            debug_file.write(debug)
            debug_file.write("\n\n")

    print("Done comparing logs. Results saved to:")
    print(" - comparison-results-line-per-line.txt")
    print(" - debug-output-line-per-line.txt")