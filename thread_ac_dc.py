import pyvisa as visa
import threading
import time

# Replace with the VISA resource string for your multimeter
VISA_ADDRESS = 'TCPIP0::192.168.222.207::INSTR'

# Function to configure and read data from the multimeter
def read_multimeter(channel, measurement_type, result):
    try:
        with visa.ResourceManager() as rm:
            multimeter = rm.open_resource(VISA_ADDRESS)
            multimeter.write(f"CONF:{1}:{measurement_type} ON,OFF,OFF")
            time.sleep(1)  # Let the multimeter settle
            measurement = multimeter.query(f"READ? {channel}")
            result.append(f"Channel {channel} {measurement_type} Measurement: {measurement}")
    except visa.VisaIOError as e:
        result.append(f"Error on Channel {channel}: {str(e)}")

# Create empty lists to store the results
dc_results = []
print(dc_results)
ac_results = []

# Create threads for DC and AC measurements
dc_thread = threading.Thread(target=read_multimeter, args=(1, 'DCV', dc_results))
ac_thread = threading.Thread(target=read_multimeter, args=(2, 'ACV', ac_results))

# Start the threads
dc_thread.start()
ac_thread.start()

# Wait for both threads to finish
dc_thread.join()
ac_thread.join()

# Print the results
for result in dc_results + ac_results:
    print(result)
