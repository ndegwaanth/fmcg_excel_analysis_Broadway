# base Docker image that we will build on
FROM python:3.12.2-slim

# set up our image by installing prerequisites; pandas in this case
RUN pip install pandas pyarrow

# set up the working directory inside the container
WORKDIR /fmcg_excel_analysis_Broadway
# copy the script to the container. 1st name is source file, 2nd is destination
COPY automated_excel_report_generator.py  automated_excel_report_generator.py

# define what to do first when the container runs
# in this example, we will just run the script
ENTRYPOINT ["python", "pipeline.py"]
