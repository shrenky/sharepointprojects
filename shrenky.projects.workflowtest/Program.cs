using System;
using System.Linq;
using System.Activities;
using System.Activities.Statements;
using System.Collections.Generic;

namespace shrenky.projects.workflowtest
{

    class Program
    {
        static void Main(string[] args)
        {
            //Activity workflow1 = new Workflow1();
            //WorkflowInvoker.Invoke(workflow1);
            Dictionary<String, Object> arguments = new Dictionary<String, Object>();
            arguments.Add("TargetCountry", "china");
            Activity instance = new Workflow1();
            WorkflowInvoker.Invoke(instance, arguments);
        }
    }
}
