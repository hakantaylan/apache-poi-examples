/**
 *
 */
package com.poi.deneme;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.List;

@Getter
@AllArgsConstructor
public class EmployeeData {

    private final String id;
    private final String name;
    private final int regularHours;
    private final SiteVisited siteVisited;

    /**
     * Return total of hours
     *
     * @param employeeList
     * @return
     */
    public static int getTotalHours(List<EmployeeData> employeeList) {
        int tolatHours = 0;

        for (EmployeeData employeeData : employeeList) {
            tolatHours += employeeData.getRegularHours();
        }

        return tolatHours;
    }

    /* (non-Javadoc)
     * @see java.lang.Object#toString()
     */
    @Override
    public String toString() {
        return "EmployeeData [id=" + id + ", name=" + name + ", regularHours=" + regularHours + ", siteVisited="
                + siteVisited + "]";
    }

    /**
     * SiteVisited class
     * @author ajaysingh
     *
     */
    @Getter
    @AllArgsConstructor
    public static class SiteVisited {
        private final String siteName;
        private final String siteUrl;

        /* (non-Javadoc)
         * @see java.lang.Object#toString()
         */
        @Override
        public String toString() {
            return "SiteVisited [siteName=" + siteName + ", siteUrl=" + siteUrl + "]";
        }
    }

}
